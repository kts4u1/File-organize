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
from urllib.parse import parse_qs, unquote, urlparse

CHUNK = 3 * 1024 * 1024  # bytes per WebDriver call when moving file contents
COMMAND_TIMEOUT = 120  # max wait for one WebDriver command; long transfers are polled instead
TRANSFER_DEADLINE = 20 * 60  # give up a download/upload after this many seconds
STALL_LIMIT = 180  # …or when a download receives nothing for this long
LOGIN_HOSTS = ("login.microsoftonline", "login.live", "/adfs/", "login.windows", "sso.")


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
        if self.sp_window in self._handles():
            self.driver.switch_to.window(self.sp_window)
        else:
            self.driver.switch_to.new_window("window")
            self.sp_window = self.driver.current_window_handle
        try:
            self.driver.switch_to.default_content()
        except Exception:
            pass


# ---------------------------------------------------------------- SharePoint
class SharePoint:
    def __init__(self, session: BrowserSession, file: SharePointFile, report: Callable[[str], None] = print,
                 log: Callable[[str], None] = lambda m: None, login_timeout: float = 150):
        self.session = session
        self.file = file
        self.report = report
        self.log = log
        self.login_timeout = login_timeout
        self.etag: str | None = None  # ETag of the version last downloaded/uploaded
        self.info: dict = {}

    @property
    def driver(self):
        return self.session.driver

    # ---------------------------------------------------------- page
    def _on_site(self) -> bool:
        return urlparse(self.session.current_url()).netloc.lower() == urlparse(self.file.origin).netloc.lower()

    def _open_page(self):
        """Be on a SharePoint page of the file's host, logged in (SSO, or the user logs in)."""
        self.session.sharepoint_window()
        if self._on_site():
            return
        self.report("SharePoint 여는 중… (로그인 창이 나오면 회사 계정으로 로그인해 주세요)")
        self.session._get(self.file.url)
        end = time.time() + self.login_timeout
        warned = False
        while time.time() < end:
            url = self.session.current_url().lower()
            if any(w in url for w in LOGIN_HOSTS):
                if not warned:
                    self.report("SharePoint 로그인이 필요합니다. 열린 Edge 창에서 로그인해 주세요.")
                    warned = True
            elif self._on_site():
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
        """Download the file inside the logged-in page and move it to Python in chunks."""
        self.driver.execute_script(DOWNLOAD_START_JS, self.file.download_url)
        state = self._poll("__tcDlState", what, progress=True)
        if not state.get("ok") or not state.get("pk"):
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
        if len(data) != size or not data.startswith(b"PK"):
            raise SharePointError("파일 내용을 끝까지 받지 못했습니다. 다시 시도하세요.")
        return data

    def _send_file(self, data: bytes, etag: str) -> dict:
        for i, start in enumerate(range(0, len(data), CHUNK)):
            self.driver.execute_script(PUSH_JS, base64.b64encode(data[start:start + CHUNK]).decode("ascii"), i == 0)
        self.driver.execute_script(UPLOAD_START_JS, self.file.site, self.file.guid, etag)
        return self._poll("__tcUpState", "SharePoint에 올리는 중", progress=False)

    # ---------------------------------------------------------- operations
    def download(self) -> tuple[bytes, dict]:
        with self.session.lock:
            try:
                self._open_page()
                info = self._file_info()
                self.report(f"SharePoint에서 받는 중… {info.get('Name', '')}")
                t0 = time.time()
                data = self._fetch_file("SharePoint에서 받는 중")
                self.etag = info.get("ETag")
                self.info = info
                self.log(f"받기 성공 {len(data):,} bytes ({time.time() - t0:.0f}초), ETag {self.etag}")
                return data, info
            finally:
                self._done()

    def upload(self, data: bytes) -> dict:
        """Replace the SharePoint file with `data` if it has not changed since our download."""
        if not data.startswith(b"PK"):
            raise SharePointError("올릴 내용이 엑셀 파일이 아닙니다.")
        if not self.etag:
            raise SharePointError("먼저 'SharePoint에서 최신 파일 받기'로 받은 파일에 반영한 뒤 올려 주세요.", code="no_base")
        with self.session.lock:
            try:
                self._open_page()
                current = self._file_info()
                if current.get("ETag") != self.etag:
                    self.log(f"ETag 불일치: 받을 때 {self.etag}, 지금 {current.get('ETag')}")
                    code, msg = explain(412)
                    raise SharePointError(msg, 412, code)
                self.report(f"SharePoint에 올리는 중… {current.get('Name', '')} ({len(data) / 1e6:.1f} MB)")
                t0 = time.time()
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
                    raise SharePointError(msg, status, code)
                self.log(f"올리기 완료 ({time.time() - t0:.0f}초)")
                after = self._file_info()
                self.etag = after.get("ETag")
                self.info = after
                # read it back to confirm what SharePoint now holds (the upload already succeeded)
                stored = b""
                try:
                    stored = self._fetch_file("올린 파일 확인 중")
                except SharePointError as exc:
                    self.log(f"올린 뒤 확인용 받기 실패(올리기는 성공): {exc}")
                sheets = sheet_names(stored) if stored.startswith(b"PK") else []
                self.log(f"올리기 성공: ETag {self.etag}, 시트 {sheets}")
                return {"ok": True, "name": after.get("Name", ""), "etag": self.etag, "modified": after.get("TimeLastModified", ""),
                        "sheets": sheets, "size": len(stored)}
            finally:
                self._done()
