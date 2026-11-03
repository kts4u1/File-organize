"""SharePoint access through Edge on the user's own PC.

Same approach as the FSN dashboard program (kts4u1/Customer): Edge runs with
a dedicated profile so the company Windows login (SSO) applies, the file's
SharePoint page is opened, and requests are made *inside that page* with
fetch(..., {credentials: 'include'}). Nothing needs an app registration or
an API key; it can do exactly what the logged-in user can do in Edge.

Download: /_layouts/15/download.aspx?UniqueId=<GUID> (proven in the FSN program).
Upload:   SharePoint REST, POST /_api/web/GetFileById('<GUID>')/$value with
          X-HTTP-Method: PUT and a form digest from /_api/contextinfo. The
          upload is refused when the file changed since it was downloaded
          (ETag check), and SharePoint keeps the previous version.
"""

from __future__ import annotations

import base64
import io
import re
import subprocess
import sys
import threading
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable
from urllib.parse import parse_qs, quote, unquote, urlparse

CHUNK = 3 * 1024 * 1024  # bytes per WebDriver call when moving file contents
COMMAND_TIMEOUT = 120  # max wait for one WebDriver command; long transfers are polled instead
TRANSFER_DEADLINE = 20 * 60  # give up a download/upload after this many seconds
STALL_LIMIT = 180  # …or when a download receives nothing for this long
LOGIN_HOSTS = ("login.microsoftonline", "login.live", "/adfs/", "login.windows", "sso.")
# Phone photos (2026-11-03): engineers upload them to this folder (next to the workbook) with the SharePoint/Teams
# app; the program brings them in and, once applied, moves them to the "done" subfolder.
DEFAULT_PHOTO_FOLDER = "툴검교정 사진"
DONE_FOLDER = "처리됨"
PHOTO_EXTS = (".jpg", ".jpeg", ".png", ".heic", ".heif", ".webp")
MAX_PHOTOS = 40  # per check
MAX_PHOTO_BYTES = 30_000_000
# The SharePoint window only does the transfers (it holds the user's login); it stays minimized and says so.
SP_TITLE = "툴 검교정 · SharePoint 연결 창 (닫지 마세요)"
SP_NOTE = ("툴 검교정 프로그램이 이 창으로 SharePoint 파일을 받고 올립니다. 닫지 말고 그대로 두세요 "
           "(평소에는 최소화되어 있고, 프로그램을 끝내면 같이 닫힙니다).")
LABEL_JS = r"""
document.title = arguments[0];
if (!document.getElementById('tc-note')) {
  const d = document.createElement('div');
  d.id = 'tc-note';
  d.textContent = arguments[1];
  d.style.cssText = 'position:fixed;left:0;right:0;top:0;z-index:2147483647;background:#f0b400;color:#1d1600;' +
    'font:600 14px "Malgun Gothic",sans-serif;padding:8px 14px;text-align:center;box-shadow:0 2px 6px rgba(0,0,0,.25)';
  document.documentElement.appendChild(d);
}
"""


def window_closed(exc: Exception) -> bool:
    """Did a browser command fail because its window was closed (e.g. the user closed the SharePoint window)?"""
    text = f"{type(exc).__name__} {exc}".lower()
    return any(w in text for w in ("nosuchwindow", "no such window", "window already closed", "web view not found",
                                   "target frame detached"))


class SharePointError(Exception):
    def __init__(self, message: str, status: int | None = None, code: str = "error"):
        super().__init__(message)
        self.status = status
        self.code = code


@dataclass
class SharePointFile:
    """A SharePoint file identified by the link people copy from Excel/SharePoint."""

    url: str  # the page link as given (Doc.aspx?sourcedoc={GUID}...)
    origin: str  # https://share.philips.com
    site: str  # https://share.philips.com/sites/HSCSCCJ-jr-all
    guid: str  # 5b1d0df6-6a62-49fb-92e3-3cb102684f84
    name: str = ""  # file name when the link carries it (&file=...)

    @classmethod
    def parse(cls, url: str) -> "SharePointFile":
        u = urlparse(url.strip())
        q = parse_qs(u.query)
        path = u.path
        if "/_layouts/15/" not in path.lower():
            raise SharePointError("SharePoint 파일 주소 형식이 아닙니다 (…/_layouts/15/Doc.aspx?sourcedoc=… 주소를 넣어 주세요).")
        guid = (q.get("sourcedoc") or q.get("UniqueId") or q.get("uniqueid") or [""])[0]
        guid = guid.strip().strip("{}").lower()
        if not re.fullmatch(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", guid):
            raise SharePointError("주소에서 파일 ID(sourcedoc)를 찾지 못했습니다.")
        site_path = path[: path.lower().index("/_layouts/15/")]
        site_path = re.sub(r"^/:[a-z]+:/[a-z]", "", site_path)  # "/:x:/r/sites/…" → "/sites/…"
        origin = f"{u.scheme}://{u.netloc}"
        name = unquote((q.get("file") or [""])[0])
        return cls(url=url.strip(), origin=origin, site=origin + site_path, guid=guid, name=name)

    @property
    def download_url(self) -> str:
        return f"{self.site}/_layouts/15/download.aspx?UniqueId={self.guid}"

    @property
    def view_url(self) -> str:
        """The file link opened in Excel for the web for viewing only (no edit lock)."""
        u = urlparse(self.url)
        if re.search(r"(^|&)action=", u.query, re.I):
            query = re.sub(r"(^|&)action=[^&]*", r"\1action=view", u.query, flags=re.I)
        else:
            query = "action=view&" + u.query if u.query else "action=view"
        return u._replace(query=query).geturl()

    @property
    def landing_url(self) -> str:
        """A page of the site that does not open the workbook ("사이트 콘텐츠").

        The file link itself (Doc.aspx?action=edit) opens Excel for the web in
        edit mode, which locks the file in the user's name, so the program's
        own upload was refused with 423 (2026-11-02, "locked for shared use by"
        the user).
        """
        return f"{self.site}/_layouts/15/viewlsts.aspx"


# ---------------------------------------------------------------- page scripts
INFO_JS = r"""
const [site, guid] = arguments; const done = arguments[arguments.length - 1];
fetch(`${site}/_api/web/GetFileById('${guid}')?$select=Name,ServerRelativeUrl,ETag,TimeLastModified,Length`,
      {credentials: 'include', headers: {Accept: 'application/json;odata=nometadata'}})
  .then(async r => done({ok: r.ok, status: r.status, data: r.ok ? await r.json() : null,
                         text: r.ok ? '' : (await r.text()).slice(0, 400)}))
  .catch(e => done({ok: false, status: 0, error: String(e)}));
"""

# Transfers run in the page in the background; Python polls their state with
# short calls, so no single WebDriver command has to outlast a big file
# (2026-11-02: a 30 s command limit cut off the first real download).
DOWNLOAD_START_JS = r"""
const url = arguments[0];
window.__tcDl = null;
const st = window.__tcDlState = {done: false, received: 0, total: 0, status: 0};
(async () => {
  try {
    const r = await fetch(url, {credentials: 'include', redirect: 'follow', cache: 'no-store'});
    st.status = r.status; st.ok = r.ok; st.cd = r.headers.get('content-disposition') || '';
    st.type = r.headers.get('content-type') || ''; st.total = Number(r.headers.get('content-length') || 0);
    const chunks = [];
    if (r.body && r.body.getReader) {
      const reader = r.body.getReader();
      for (;;) {
        const {done, value} = await reader.read();
        if (done) break;
        chunks.push(value); st.received += value.length;
      }
    } else {
      const b = new Uint8Array(await r.arrayBuffer()); chunks.push(b); st.received = b.length;
    }
    const buf = new Uint8Array(st.received); let o = 0;
    for (const c of chunks) { buf.set(c, o); o += c.length; }
    window.__tcDl = buf; st.size = buf.length;
    st.pk = buf.length > 1 && buf[0] === 0x50 && buf[1] === 0x4b;
  } catch (e) { st.error = String(e); }
  st.done = true;
})();
return true;
"""

STATE_JS = "return window[arguments[0]] || null;"

# Shared by the folder scripts: '@a' parameter alias (paths may hold spaces, Korean, quotes) and the form digest.
_FOLDER_HELPERS = r"""
const alias = p => '@a=' + encodeURIComponent("'" + p.replace(/'/g, "''") + "'");
const NOMETA = 'application/json;odata=nometadata';
const digestOf = async site => {
  const r = await fetch(`${site}/_api/contextinfo`, {method: 'POST', credentials: 'include', headers: {Accept: NOMETA}});
  if (!r.ok) throw Object.assign(new Error('contextinfo'), {status: r.status, text: (await r.text()).slice(0, 300)});
  const j = await r.json();
  return j.FormDigestValue || (j.d && j.d.GetContextWebInformation && j.d.GetContextWebInformation.FormDigestValue);
};
const folderExists = async (site, path) => {
  const r = await fetch(`${site}/_api/web/GetFolderByServerRelativePath(decodedurl=@a)?$select=Exists&${alias(path)}`,
                        {credentials: 'include', headers: {Accept: NOMETA}});
  if (r.status === 404) return false;
  if (!r.ok) throw Object.assign(new Error('folder'), {status: r.status, text: (await r.text()).slice(0, 300)});
  return !!(await r.json()).Exists;
};
const makeFolder = async (site, path, digest) => {
  const opts = {method: 'POST', credentials: 'include', headers: {Accept: NOMETA, 'X-RequestDigest': digest}};
  let r = await fetch(`${site}/_api/web/folders/AddUsingPath(decodedurl=@a)?${alias(path)}`, opts);
  if (!r.ok) r = await fetch(`${site}/_api/web/folders/add(url=@a)?${alias(path)}`, opts);  // older servers
  if (!r.ok) throw Object.assign(new Error('create folder'), {status: r.status, text: (await r.text()).slice(0, 300)});
};
"""

# List the photo folder (creating it when asked and missing).
PHOTOS_JS = _FOLDER_HELPERS + r"""
const [site, path, create] = arguments; const done = arguments[arguments.length - 1];
(async () => {
  try {
    let created = false;
    if (!(await folderExists(site, path))) {
      if (!create) return done({ok: true, exists: false, created, files: []});
      await makeFolder(site, path, await digestOf(site));
      created = true;
    }
    const r = await fetch(`${site}/_api/web/GetFolderByServerRelativePath(decodedurl=@a)/Files` +
                          `?$select=Name,UniqueId,Length,TimeCreated,ServerRelativeUrl,Author/Title&$expand=Author&$top=500&${alias(path)}`,
                          {credentials: 'include', headers: {Accept: NOMETA}});
    if (!r.ok) return done({ok: false, status: r.status, step: 'list', text: (await r.text()).slice(0, 300)});
    const j = await r.json();
    const files = (j.value || []).map(f => ({id: String(f.UniqueId || '').toLowerCase(), name: f.Name, size: Number(f.Length || 0),
      created: f.TimeCreated || '', url: f.ServerRelativeUrl || `${path}/${f.Name}`, author: (f.Author && f.Author.Title) || ''}));
    done({ok: true, exists: true, created, files});
  } catch (e) { done({ok: false, status: e.status || 0, step: e.message, text: e.text || String(e)}); }
})();
"""

# Move processed photos into the done folder (keeping both when a name is taken).
MOVE_JS = _FOLDER_HELPERS + r"""
const [site, origin, doneFolder, items] = arguments; const done = arguments[arguments.length - 1];
(async () => {
  const results = [];
  try {
    const digest = await digestOf(site);
    if (!(await folderExists(site, doneFolder))) await makeFolder(site, doneFolder, digest);
    for (const it of items) {
      const verbose = 'application/json;odata=verbose';
      let r = await fetch(`${site}/_api/SP.MoveCopyUtil.MoveFileByPath()`, {method: 'POST', credentials: 'include',
        headers: {Accept: verbose, 'Content-Type': verbose, 'X-RequestDigest': digest},
        body: JSON.stringify({
          srcPath: {__metadata: {type: 'SP.ResourcePath'}, DecodedUrl: origin + it.from},
          destPath: {__metadata: {type: 'SP.ResourcePath'}, DecodedUrl: origin + it.to},
          options: {__metadata: {type: 'SP.MoveCopyOptions'}, KeepBoth: true, ResetAuthorAndCreatedOnCopy: false,
                    ShouldBypassSharedLocks: true}})});
      if (!r.ok && (r.status === 400 || r.status === 404 || r.status === 405)) {
        // older servers: move by id under a unique name
        r = await fetch(`${site}/_api/web/GetFileById('${it.id}')/moveto(newurl=@a,flags=1)?${alias(it.alt)}`,
                        {method: 'POST', credentials: 'include', headers: {Accept: NOMETA, 'X-RequestDigest': digest}});
      }
      results.push({id: it.id, ok: r.ok, status: r.status, text: r.ok ? '' : (await r.text()).slice(0, 300)});
    }
    done({ok: true, results});
  } catch (e) { done({ok: false, status: e.status || 0, step: e.message, text: e.text || String(e), results}); }
})();
"""

CHUNK_GET_JS = r"""
const buf = window.__tcDl; if (!buf) return null;
const part = buf.subarray(arguments[0], arguments[0] + arguments[1]);
let bin = '';
for (let i = 0; i < part.length; i += 0x8000) bin += String.fromCharCode.apply(null, part.subarray(i, i + 0x8000));
if (arguments[0] + arguments[1] >= buf.length) window.__tcDl = null;
return btoa(bin);
"""

PUSH_JS = r"""
if (arguments[1]) window.__tcUp = [];
window.__tcUp.push(arguments[0]);
return window.__tcUp.length;
"""

UPLOAD_START_JS = r"""
const [site, guid, etag] = arguments;
const st = window.__tcUpState = {done: false, status: 0, step: 'contextinfo'};
(async () => {
  try {
    const ctx = await fetch(`${site}/_api/contextinfo`, {method: 'POST', credentials: 'include',
                            headers: {Accept: 'application/json;odata=nometadata'}});
    if (!ctx.ok) {
      Object.assign(st, {ok: false, status: ctx.status, text: (await ctx.text()).slice(0, 400)});
    } else {
      const j = await ctx.json();
      const digest = j.FormDigestValue || (j.d && j.d.GetContextWebInformation && j.d.GetContextWebInformation.FormDigestValue);
      const parts = (window.__tcUp || []).map(b64 => {
        const bin = atob(b64); const u = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
        return u;
      });
      window.__tcUp = null;
      const headers = {'X-RequestDigest': digest, 'X-HTTP-Method': 'PUT'};
      if (etag) headers['If-Match'] = etag;
      st.step = 'upload';
      const r = await fetch(`${site}/_api/web/GetFileById('${guid}')/$value`,
                            {method: 'POST', credentials: 'include', headers, body: new Blob(parts)});
      Object.assign(st, {ok: r.ok, status: r.status, text: r.ok ? '' : (await r.text()).slice(0, 600)});
    }
  } catch (e) { Object.assign(st, {ok: false, status: 0, error: String(e)}); }
  st.done = true;
})();
return true;
"""


def sheet_names(data: bytes) -> list[str]:
    """Sheet names of an .xlsx, read straight from xl/workbook.xml."""
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        xml = zf.read("xl/workbook.xml").decode("utf-8", "replace")
    names = re.findall(r'<(?:\w+:)?sheet\b[^>]*\bname="([^"]*)"', xml)
    return [n.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"') for n in names]


def lock_owner(text: str) -> str:
    """Who holds a SharePoint lock, from a 423 answer ('…locked for shared use by a.b@philips.com [membership]')."""
    m = re.search(r"locked for (?:shared|exclusive) use by ([^\s<\[]+)", text or "")
    return m.group(1) if m else ""


def explain(status: int | None, step: str = "") -> tuple[str, str]:
    """(code, Korean message) for a SharePoint HTTP status."""
    if status == 423:
        return "locked", ("파일이 잠겨 있습니다. 누군가(또는 다른 창의 본인이) Excel에서 이 파일을 열어 편집 중입니다. "
                          "모두 닫은 뒤 다시 올리거나, 'Excel 웹 스크립트' 방식을 쓰세요.")
    if status == 412:
        return "changed", "받은 뒤에 SharePoint 파일이 바뀌었습니다. 최신 파일을 다시 받아서 다시 반영하세요."
    if status in (401, 403):
        return "denied", "SharePoint 권한이 없거나 로그인이 풀렸습니다. 열린 Edge 창에서 로그인했는지, 이 파일의 편집 권한이 있는지 확인하세요."
    if status == 404:
        return "missing", "SharePoint에서 파일을 찾지 못했습니다. 주소(config.json의 sharepoint_url)를 확인하세요."
    if status == 0:
        return "network", "SharePoint에 연결하지 못했습니다. 네트워크/VPN을 확인하세요."
    return "error", f"SharePoint 오류 ({step} {status})."


# ---------------------------------------------------------------- browser
class BrowserSession:
    """One Selenium-driven browser: a window for the tool page and one for SharePoint.

    Edge with a dedicated profile on Windows (company SSO). Tests can use
    Chrome/Chromium with `browser="chrome"`.
    """

    def __init__(self, profile_dir: Path, report: Callable[[str], None] = print, browser: str = "edge",
                 binary: str | None = None, headless: bool = False, log: Callable[[str], None] = lambda m: None):
        self.profile_dir = Path(profile_dir)
        self.report = report
        self.log = log
        self.browser = browser
        self.binary = binary
        self.headless = headless
        self.driver = None
        self.tool_window: str | None = None
        self.sp_window: str | None = None
        self.lock = threading.RLock()

    def start(self):
        from selenium import webdriver

        if self.browser == "chrome":
            from selenium.webdriver.chrome.options import Options
            from selenium.webdriver.chrome.service import Service
        else:
            from selenium.webdriver.edge.options import Options
            from selenium.webdriver.edge.service import Service

        def options(profile: Path):
            opts = Options()
            opts.page_load_strategy = "eager"
            opts.add_argument(f"--user-data-dir={profile}")
            opts.add_argument("--profile-directory=Default")
            for arg in ("--no-first-run", "--no-default-browser-check", "--start-maximized"):
                opts.add_argument(arg)
            if self.headless:
                for arg in ("--headless=new", "--no-sandbox", "--disable-dev-shm-usage", "--window-size=1400,1000"):
                    opts.add_argument(arg)
            if self.binary:
                opts.binary_location = self.binary
            opts.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
            return opts

        def service():
            s = Service()
            if sys.platform == "win32":
                try:
                    s.creation_flags = subprocess.CREATE_NO_WINDOW  # no console window
                except Exception:
                    pass
            return s

        factory = webdriver.Chrome if self.browser == "chrome" else webdriver.Edge
        self.profile_dir.mkdir(parents=True, exist_ok=True)
        self.report("Edge 브라우저를 여는 중…")
        try:
            self.driver = factory(options=options(self.profile_dir), service=service())
        except Exception as exc:
            # a previous run left the profile locked → try a second profile once
            self.log(f"브라우저 시작 실패 → 다른 프로필로 재시도: {exc}")
            alt = self.profile_dir.with_name(self.profile_dir.name + "_2")
            alt.mkdir(parents=True, exist_ok=True)
            self.driver = factory(options=options(alt), service=service())
        self.driver.set_page_load_timeout(45)
        try:
            self.driver.command_executor._client_config.timeout = COMMAND_TIMEOUT  # selenium 4.2x+
        except Exception:
            try:
                self.driver.command_executor.set_timeout(COMMAND_TIMEOUT)
            except Exception:
                pass

    def open_tool(self, url: str):
        with self.lock:
            if self.driver is None:
                self.start()
            self._get(url)
            self.tool_window = self.driver.current_window_handle

    def show_tool(self):
        with self.lock:
            if self.driver is not None and self.tool_window in self._handles():
                self.driver.switch_to.window(self.tool_window)

    def alive(self) -> bool:
        try:
            return self.driver is not None and bool(self.driver.window_handles)
        except Exception:
            return False

    def tool_open(self) -> bool:
        """Is the tool window still open? (Waits while a download/upload holds the browser.)"""
        with self.lock:
            return self.driver is not None and self.tool_window is not None and self.tool_window in self._handles()

    def quit(self):
        with self.lock:
            if self.driver is not None:
                try:
                    self.driver.quit()
                except Exception:
                    pass
                self.driver = None

    def _handles(self) -> list[str]:
        try:
            return list(self.driver.window_handles)
        except Exception:
            return []

    def _get(self, url: str):
        try:
            self.driver.get(url)
        except Exception as exc:  # eager load + timeout: keep going like the FSN program
            self.log(f"driver.get 예외(무시): {type(exc).__name__}: {str(exc)[:200]}")

    def current_url(self) -> str:
        try:
            return self.driver.current_url or ""
        except Exception:
            return ""

    def sharepoint_window(self):
        """Switch to the SharePoint window, creating it when needed."""
        handles = self._handles()
        if self.sp_window in handles:
            self.driver.switch_to.window(self.sp_window)
        else:
            # the window we were on may be the one the user just closed: start from one that is open
            start = self.tool_window if self.tool_window in handles else (handles[0] if handles else None)
            if start:
                self.driver.switch_to.window(start)
            self.driver.switch_to.new_window("window")
            self.sp_window = self.driver.current_window_handle
        try:
            self.driver.switch_to.default_content()
        except Exception:
            pass

    def tuck_sharepoint(self):
        """Minimize the SharePoint window (current window): the user works in the tool window."""
        try:
            self.driver.minimize_window()
        except Exception as exc:
            self.log(f"SharePoint 창 최소화 실패(무시): {type(exc).__name__}: {str(exc)[:120]}")

    def show_sharepoint(self):
        """Bring the SharePoint window (current window) up, e.g. for the company login."""
        try:
            self.driver.maximize_window()
            self.driver.switch_to.window(self.sp_window)
        except Exception as exc:
            self.log(f"SharePoint 창 보이기 실패(무시): {type(exc).__name__}: {str(exc)[:120]}")

    def label_sharepoint(self):
        try:
            self.driver.execute_script(LABEL_JS, SP_TITLE, SP_NOTE)
        except Exception as exc:
            self.log(f"SharePoint 창 안내 표시 실패(무시): {type(exc).__name__}: {str(exc)[:120]}")


# ---------------------------------------------------------------- SharePoint
class SharePoint:
    def __init__(self, session: BrowserSession, file: SharePointFile, report: Callable[[str], None] = print,
                 log: Callable[[str], None] = lambda m: None, login_timeout: float = 150, photo_folder: str | None = None):
        self.session = session
        self.file = file
        self.photo_folder = (photo_folder or DEFAULT_PHOTO_FOLDER).strip().rstrip("/")
        self.photos: dict[str, dict] = {}  # phone photos seen in the folder, by UniqueId
        self.report = report
        self.log = log
        self.login_timeout = login_timeout
        self.etag: str | None = None  # ETag of the version last downloaded/uploaded
        self.info: dict = {}
        self.last_ok = 0.0  # when the SharePoint page last worked; reloaded after a while (login, sleeping tabs)
        self.sending = False  # the upload request may have reached SharePoint
        self.keep_front = False  # leave the SharePoint window in front after this operation

    @property
    def driver(self):
        return self.session.driver

    # ---------------------------------------------------------- page
    def _on_site(self) -> bool:
        return urlparse(self.session.current_url()).netloc.lower() == urlparse(self.file.origin).netloc.lower()

    def _on_workbook(self) -> bool:
        """Is the SharePoint window showing the workbook in Excel (which holds a lock)?"""
        path = urlparse(self.session.current_url()).path.lower()
        return path.endswith("/doc.aspx") or path.endswith("/doc2.aspx") or "/_layouts/15/xl" in path

    def _open_page(self):
        """Be on a SharePoint page of the file's host, logged in (SSO, or the user logs in).

        The SharePoint window stays minimized; it comes up only while a login page waits for the user.
        """
        self.session.sharepoint_window()
        self.session.tuck_sharepoint()
        if self._on_site() and not self._on_workbook() and time.time() - self.last_ok < 600:
            return
        self.report("SharePoint 여는 중…")
        self.session._get(self.file.landing_url)
        end = time.time() + self.login_timeout
        login_since = None
        shown = False
        while time.time() < end:
            url = self.session.current_url().lower()
            if any(w in url for w in LOGIN_HOSTS):
                login_since = login_since or time.time()
                # company SSO usually passes by itself in a moment; show the window only if it waits for the user
                if not shown and time.time() - login_since >= 3:
                    self.session.show_sharepoint()
                    self.report("SharePoint 로그인이 필요합니다. 앞에 뜬 SharePoint 창에서 회사 계정으로 로그인해 주세요.")
                    shown = True
            elif self._on_site():
                self.session.label_sharepoint()
                if shown:
                    self.session.tuck_sharepoint()
                self.last_ok = time.time()
                return
            time.sleep(1)
        raise SharePointError("SharePoint에 로그인하지 못했습니다 (시간 초과).", code="login")

    def _async(self, script: str, *args, timeout: float = 300):
        self.driver.set_script_timeout(timeout)
        return self.driver.execute_async_script(script, *args)

    def _file_info(self) -> dict:
        res = self._async(INFO_JS, self.file.site, self.file.guid, timeout=60)
        if not isinstance(res, dict) or not res.get("ok"):
            status = res.get("status") if isinstance(res, dict) else None
            code, msg = explain(status, "파일 정보")
            self.log(f"파일 정보 실패: {res}")
            raise SharePointError(msg, status, code)
        return res["data"] or {}

    def _done(self):
        if self.keep_front:  # the uploaded file is on show in the SharePoint window: leave it in front
            self.keep_front = False
            return
        self.session.show_tool()

    def _poll(self, key: str, what: str, progress: bool) -> dict:
        """Wait for a background transfer in the page, reporting progress."""
        start = last_change = last_report = time.time()
        last = -1
        while True:
            state = self.driver.execute_script(STATE_JS, key) or {}
            if state.get("done"):
                return state
            now = time.time()
            received = int(state.get("received") or 0)
            if received != last:
                last, last_change = received, now
            if now - last_report >= 3:
                last_report = now
                if progress and received:
                    total = int(state.get("total") or 0)
                    size = f"{received / 1e6:.1f} / {total / 1e6:.1f} MB" if total else f"{received / 1e6:.1f} MB"
                    self.report(f"{what}… {size}")
                elif now - start >= 10:
                    self.report(f"{what}… ({int(now - start)}초)")
            if progress and now - last_change > STALL_LIMIT:
                raise SharePointError(f"{what}이(가) {STALL_LIMIT // 60}분 동안 진행되지 않아 멈췄습니다. 네트워크를 확인하고 다시 시도하세요.",
                                      code="stalled")
            if now - start > TRANSFER_DEADLINE:
                raise SharePointError(f"{what} 시간이 너무 오래 걸려 멈췄습니다 ({TRANSFER_DEADLINE // 60}분).", code="timeout")
            time.sleep(1)

    def _fetch_file(self, what: str) -> bytes:
        """Download the workbook inside the logged-in page and move it to Python in chunks."""
        return self._fetch(self.file.download_url, what)

    def _fetch(self, url: str, what: str, photo: bool = False) -> bytes:
        self.driver.execute_script(DOWNLOAD_START_JS, url)
        state = self._poll("__tcDlState", what, progress=not photo)
        if photo:
            if not state.get("ok") or "text/html" in str(state.get("type", "")):
                self.log(f"{what} 실패: { {k: v for k, v in state.items() if k != 'cd'} }")
                code, msg = explain(state.get("status") if not state.get("ok") else 403, "사진 받기")
                raise SharePointError(msg, state.get("status"), code)
        elif not state.get("ok") or not state.get("pk"):
            status = state.get("status")
            self.log(f"{what} 실패: { {k: v for k, v in state.items() if k != 'cd'} }")
            if status == 200:  # an HTML page instead of the workbook: usually the login expired
                raise SharePointError("엑셀 대신 다른 페이지가 받아졌습니다. Edge 창에서 SharePoint 로그인 상태를 확인하고 다시 시도하세요.",
                                      status, "denied")
            code, msg = explain(status, "받기")
            raise SharePointError(msg, status, code)
        size = int(state["size"])
        parts = []
        for start in range(0, size, CHUNK):
            b64 = self.driver.execute_script(CHUNK_GET_JS, start, CHUNK)
            parts.append(base64.b64decode(b64 or ""))
        data = b"".join(parts)
        if len(data) != size or not (photo or data.startswith(b"PK")):
            raise SharePointError("파일 내용을 끝까지 받지 못했습니다. 다시 시도하세요.")
        return data

    def _send_file(self, data: bytes, etag: str) -> dict:
        for i, start in enumerate(range(0, len(data), CHUNK)):
            self.driver.execute_script(PUSH_JS, base64.b64encode(data[start:start + CHUNK]).decode("ascii"), i == 0)
        self.driver.execute_script(UPLOAD_START_JS, self.file.site, self.file.guid, etag)
        return self._poll("__tcUpState", "SharePoint에 올리는 중", progress=False)

    # ---------------------------------------------------------- operations
    def _again_if_window_closed(self, what: str, run):
        """Run `run()`; if the SharePoint window was closed meanwhile, open it again and retry once."""
        with self.session.lock:
            for attempt in (1, 2):
                self.sending = False
                try:
                    return run()
                except Exception as exc:
                    if not window_closed(exc):
                        raise
                    self.log(f"{what} 중 SharePoint 창이 닫힘 ({attempt}회째): {type(exc).__name__}: {str(exc)[:160]}")
                    if self.sending:
                        raise SharePointError("올리는 중에 SharePoint 창이 닫혀 중단됐습니다. '다시 올리기'를 눌러 주세요 "
                                              "(이미 올라갔다면 '최신 파일로 다시 반영'이 나옵니다).", code="closed") from exc
                    if attempt == 2:
                        raise SharePointError("SharePoint 창이 닫혀 진행하지 못했습니다. 다시 눌러 주세요.", code="closed") from exc
                    self.report("SharePoint 창이 닫혀 다시 여는 중…")
                finally:
                    self._done()

    def download(self) -> tuple[bytes, dict]:
        return self._again_if_window_closed("받기", self._download)

    def _download(self) -> tuple[bytes, dict]:
        self._open_page()
        info = self._file_info()
        self.report(f"SharePoint에서 받는 중… {info.get('Name', '')}")
        t0 = time.time()
        data = self._fetch_file("SharePoint에서 받는 중")
        self.etag = info.get("ETag")
        self.info = info
        self.log(f"받기 성공 {len(data):,} bytes ({time.time() - t0:.0f}초), ETag {self.etag}")
        return data, info

    # ---------------------------------------------------------- phone photos
    def photo_folder_path(self) -> str:
        """Server-relative path of the phone-photo folder: a name next to the workbook, or a full '/sites/…' path."""
        if self.photo_folder.startswith("/"):
            return self.photo_folder
        if not self.info.get("ServerRelativeUrl"):
            self.info = self._file_info()
        parent = str(self.info.get("ServerRelativeUrl", "")).rsplit("/", 1)[0]
        return f"{parent}/{self.photo_folder}"

    def folder_link(self, path: str) -> str:
        """Web address of a folder, for the phone (opens in the SharePoint app or the browser)."""
        return self.file.origin + quote(path)

    def list_photos(self) -> dict:
        """New phone photos in the folder (the folder is created when it does not exist yet)."""
        def run():
            self._open_page()
            path = self.photo_folder_path()
            res = self._async(PHOTOS_JS, self.file.site, path, True, timeout=60)
            if not isinstance(res, dict) or not res.get("ok"):
                self.log(f"사진 폴더 확인 실패: {res}")
                status = res.get("status") if isinstance(res, dict) else None
                code, msg = explain(status, "사진 폴더")
                if status == 404:
                    msg = "SharePoint 사진 폴더를 찾지 못했습니다. config.json의 photo_folder를 확인하세요."
                raise SharePointError(msg, status, code)
            if res.get("created"):
                self.report(f"SharePoint에 휴대폰 사진 폴더를 만들었습니다: {self.photo_folder}")
            photos, skipped = [], 0
            for f in sorted(res.get("files") or [], key=lambda f: f.get("created", "")):
                name = str(f.get("name", ""))
                if not name.lower().endswith(PHOTO_EXTS):
                    continue
                if int(f.get("size") or 0) > MAX_PHOTO_BYTES or len(photos) >= MAX_PHOTOS:
                    skipped += 1
                    continue
                self.photos[f["id"]] = f
                photos.append(f)
            if photos:
                self.report(f"휴대폰 사진 폴더에 새 사진 {len(photos)}장")
            return {"ok": True, "folder": path, "link": self.folder_link(path), "name": self.photo_folder,
                    "created": bool(res.get("created")), "photos": photos, "skipped": skipped}
        return self._again_if_window_closed("사진 폴더 확인", run)

    def fetch_photo(self, uid: str) -> tuple[bytes, dict]:
        uid = (uid or "").strip().strip("{}").lower()
        meta = self.photos.get(uid)
        if meta is None:
            raise SharePointError("목록에 없는 사진입니다. 휴대폰 사진을 다시 확인해 주세요.", code="missing")

        def run():
            self._open_page()
            url = f"{self.file.site}/_layouts/15/download.aspx?UniqueId={uid}"
            return self._fetch(url, f"사진 받는 중 {meta.get('name', '')}", photo=True), meta
        return self._again_if_window_closed("사진 받기", run)

    def _archive_photos(self, ids) -> dict:
        """Move applied phone photos into the done folder (after a successful upload). Failures are reported, not raised."""
        items = []
        stamp = time.strftime("%Y%m%d-%H%M%S")
        for uid in dict.fromkeys(str(i).strip().lower() for i in ids or []):
            meta = self.photos.get(uid)
            if not meta:
                continue
            folder, name = str(meta["url"]).rsplit("/", 1)
            items.append({"id": uid, "from": meta["url"], "to": f"{folder}/{DONE_FOLDER}/{name}",
                          "alt": f"{folder}/{DONE_FOLDER}/{stamp}_{name}"})
        if not items:
            return {"moved": 0, "failed": 0}
        done_folder = items[0]["to"].rsplit("/", 1)[0]
        try:
            res = self._async(MOVE_JS, self.file.site, self.file.origin, done_folder, items, timeout=120)
        except Exception as exc:
            if window_closed(exc):
                raise
            self.log(f"사진 옮기기 예외: {type(exc).__name__}: {str(exc)[:200]}")
            return {"moved": 0, "failed": len(items), "error": "사진을 '처리됨' 폴더로 옮기지 못했습니다."}
        results = (res or {}).get("results") or []
        moved = [r["id"] for r in results if r.get("ok")]
        for uid in moved:
            self.photos.pop(uid, None)
        failed = len(items) - len(moved)
        if failed or not (res or {}).get("ok"):
            self.log(f"사진 옮기기 일부 실패: {res}")
        if moved:
            self.report(f"반영한 휴대폰 사진 {len(moved)}장을 '{DONE_FOLDER}' 폴더로 옮겼습니다.")
        out = {"moved": len(moved), "failed": failed}
        if failed:
            out["error"] = f"사진 {failed}장을 '{DONE_FOLDER}' 폴더로 옮기지 못했습니다. 다음에 다시 나올 수 있습니다."
        return out

    def upload(self, data: bytes, done_photos=()) -> dict:
        """Replace the SharePoint file with `data` if it has not changed since our download."""
        if not data.startswith(b"PK"):
            raise SharePointError("올릴 내용이 엑셀 파일이 아닙니다.")
        if not self.etag:
            raise SharePointError("먼저 'SharePoint에서 최신 파일 받기'로 받은 파일에 반영한 뒤 올려 주세요.", code="no_base")
        return self._again_if_window_closed("올리기", lambda: self._upload(data, done_photos))

    def _upload(self, data: bytes, done_photos=()) -> dict:
        self._open_page()
        current = self._file_info()
        if current.get("ETag") != self.etag:
            self.log(f"ETag 불일치: 받을 때 {self.etag}, 지금 {current.get('ETag')}")
            code, msg = explain(412)
            raise SharePointError(msg, 412, code)
        self.report(f"SharePoint에 올리는 중… {current.get('Name', '')} ({len(data) / 1e6:.1f} MB)")
        t0 = time.time()
        self.sending = True
        res = self._send_file(data, self.etag)
        if res.get("status") == 412:
            # some servers reject the file ETag in If-Match: re-check, then send without it
            again = self._file_info()
            if again.get("ETag") == self.etag:
                self.log("If-Match 거부 → ETag 재확인 후 다시 올림")
                res = self._send_file(data, "")
        if not isinstance(res, dict) or not res.get("ok"):
            status = res.get("status") if isinstance(res, dict) else None
            self.log(f"올리기 실패: {res}")
            code, msg = explain(status, res.get("step", "") if isinstance(res, dict) else "")
            owner = lock_owner(res.get("text", "") if isinstance(res, dict) else "")
            if status == 423 and owner:
                msg = (f"{owner} 님이 이 파일을 Excel에서 열어 두어 SharePoint가 잠가 두었습니다. "
                       "누군가 열어 둔 동안에는 파일째 바꿀 수 없으니, 아래 'Excel 웹 스크립트로 반영'을 쓰세요.")
            raise SharePointError(msg, status, code)
        self.log(f"올리기 완료 ({time.time() - t0:.0f}초)")
        self.sending = False
        # the upload succeeded: what follows only confirms it, so its failures (e.g. a closed window) are not errors
        after: dict = {}
        stored = b""
        try:
            after = self._file_info()
            stored = self._fetch_file("올린 파일 확인 중")  # read back what SharePoint now holds
        except Exception as exc:
            self.log(f"올린 뒤 확인 실패(올리기는 성공): {type(exc).__name__}: {str(exc)[:160]}")
        self.etag = after.get("ETag")  # unknown → the next upload asks for a fresh download first
        self.info = after
        sheets = sheet_names(stored) if stored.startswith(b"PK") else []
        self.log(f"올리기 성공: ETag {self.etag}, 시트 {sheets}")
        photos = {"moved": 0, "failed": 0}
        try:
            photos = self._archive_photos(done_photos)
        except Exception as exc:  # the upload already succeeded
            self.log(f"사진 옮기기 실패(올리기는 성공): {type(exc).__name__}: {str(exc)[:160]}")
            photos = {"moved": 0, "failed": len(done_photos), "error": "사진을 '처리됨' 폴더로 옮기지 못했습니다."}
        opened = self._show_workbook()
        return {"ok": True, "name": after.get("Name", "") or self.file.name, "etag": self.etag,
                "modified": after.get("TimeLastModified", ""), "sheets": sheets, "size": len(stored), "opened": opened,
                "photos": photos}

    def _show_workbook(self) -> bool:
        """After an upload, show the file in Excel for the web so the user can check it right away (2026-11-02 request).

        View mode: edit mode would lock the file in the user's name and refuse the next upload (423). The next
        download/upload leaves this page first (`_open_page`).
        """
        try:
            self.report("올린 파일을 SharePoint 창에 엽니다 (보기 모드)…")
            self.session._get(self.file.view_url)
            self.session.show_sharepoint()
            self.keep_front = True
            return True
        except Exception as exc:
            self.log(f"올린 파일 열기 실패(무시): {type(exc).__name__}: {str(exc)[:160]}")
            return False
