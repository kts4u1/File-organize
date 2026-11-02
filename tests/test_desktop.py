import json
import os
import shutil
import urllib.error
import urllib.request

import pytest

from conftest import build_workbook
from desktop.server import CDN_TO_LOCAL, ToolServer
from desktop.sharepoint import SharePointError, SharePointFile, explain, sheet_names
from mock_sharepoint import GUID, MockSharePoint

USER_URL = ("https://share.philips.com/:x:/r/sites/HSCSCCJ-jr-all/_layouts/15/Doc.aspx?action=edit"
            "&sourcedoc=%7B5b1d0df6-6a62-49fb-92e3-3cb102684f84%7D&wdExp=TEAMS-TREATMENT&web=1")
WEB = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "web")


def test_parse_user_link():
    f = SharePointFile.parse(USER_URL)
    assert f.origin == "https://share.philips.com"
    assert f.site == "https://share.philips.com/sites/HSCSCCJ-jr-all"
    assert f.guid == "5b1d0df6-6a62-49fb-92e3-3cb102684f84"
    assert f.download_url == ("https://share.philips.com/sites/HSCSCCJ-jr-all/_layouts/15/download.aspx"
                              "?UniqueId=5b1d0df6-6a62-49fb-92e3-3cb102684f84")


def test_parse_fco_style_link_with_name():
    f = SharePointFile.parse("https://share.philips.com/:x:/r/sites/KRHSFCOstatus/_layouts/15/Doc.aspx?sourcedoc="
                             "%7B41877D54-C06C-43EA-A61A-D8EE1F661BF5%7D&file=MR%20FCO%20Action%20Plan.xlsx")
    assert f.site.endswith("/sites/KRHSFCOstatus") and f.guid == "41877d54-c06c-43ea-a61a-d8ee1f661bf5"
    assert f.name == "MR FCO Action Plan.xlsx"


def test_view_link_does_not_open_the_workbook_for_editing():
    """Edit mode locks the file in the user's name (423 on the next upload); the check after an upload uses view mode."""
    f = SharePointFile.parse(USER_URL)
    assert f.view_url == USER_URL.replace("action=edit", "action=view")
    g = SharePointFile.parse("https://share.philips.com/:x:/r/sites/K/_layouts/15/Doc.aspx?sourcedoc=%7B41877D54-C06C-43EA-"
                             "A61A-D8EE1F661BF5%7D&file=a.xlsx")
    assert g.view_url.startswith("https://share.philips.com/:x:/r/sites/K/_layouts/15/Doc.aspx?action=view&sourcedoc=")


def test_parse_rejects_other_links():
    with pytest.raises(SharePointError):
        SharePointFile.parse("https://share.philips.com/sites/HSCSCCJ-jr-all/Shared%20Documents/x.xlsx")


def test_lock_owner_from_sharepoint_answer():
    from desktop.sharepoint import lock_owner
    text = ('<m:message xml:lang="en-US">The file https://share.philips.com/sites/HSCSCCJ-jr-all/Shared Documents/'
            'Tool list for Field validation 20250325.xlsx is locked for shared use by taesu.kim@philips.com [membership].</m:message>')
    assert lock_owner(text) == "taesu.kim@philips.com"
    assert lock_owner("other error") == ""


def test_explain_codes():
    assert explain(423)[0] == "locked" and explain(412)[0] == "changed" and explain(403)[0] == "denied"


def test_sheet_names(workbook_path):
    with open(workbook_path, "rb") as fh:
        assert sheet_names(fh.read()) == ["대전", "광주"]


class FakeSharePoint:
    def __init__(self, data):
        self.data = data
        self.file = SharePointFile.parse(USER_URL)
        self.uploaded = None
        self.fail = None

    def download(self):
        return self.data, {"Name": "Tool Verification.xlsx", "TimeLastModified": "2026-11-02T01:00:00Z"}

    def upload(self, data):
        if self.fail:
            raise SharePointError("잠김", 423, "locked")
        self.uploaded = data
        return {"ok": True, "name": "Tool Verification.xlsx", "sheets": ["대전", "광주", "TEST"]}


@pytest.fixture
def server(workbook_path, tmp_path):
    with open(workbook_path, "rb") as fh:
        fake = FakeSharePoint(fh.read())
    opened = []
    srv = ToolServer(WEB, tmp_path / "out", fake, folders={"output": tmp_path / "out", "logs": tmp_path},
                     opener=opened.append).start()
    srv.opened = opened
    yield srv, fake
    srv.stop()


def _req(srv, path, method="GET", body=None, token=True):
    req = urllib.request.Request(srv.url.rstrip("/") + path, data=body, method=method)
    if token:
        req.add_header("X-TC-Token", srv.token)
    try:
        with urllib.request.urlopen(req) as r:
            return r.status, dict(r.headers), r.read()
    except urllib.error.HTTPError as e:
        return e.code, dict(e.headers), e.read()


def test_page_is_served_with_local_libraries(server):
    srv, _ = server
    status, _, body = _req(srv, "/", token=False)
    page = body.decode()
    assert status == 200 and page.startswith("<!doctype html>")
    assert f'name="tc-token" content="{srv.token}"' in page
    for cdn, local in CDN_TO_LOCAL.items():
        assert cdn not in page and local in page
        assert _req(srv, local, token=False)[0] == 200
    assert _req(srv, "/tesseract/worker-wrap.js", token=False)[0] == 200
    assert _req(srv, "/tesseract/manifest.json", token=False)[0] == 200
    assert _req(srv, "/../desktop/app.py", token=False)[0] == 404


def test_api_needs_token(server):
    srv, _ = server
    assert _req(srv, "/api/status", token=False)[0] == 403
    assert _req(srv, "/api/sp/download", method="POST", token=False)[0] == 403
    status, _, body = _req(srv, "/api/status")
    assert status == 200 and json.loads(body)["sharepoint"] is True
    assert _req(srv, "/api/progress", token=False)[0] == 403
    srv.note("SharePoint에서 받는 중… 12.0 / 207.8 MB")
    assert json.loads(_req(srv, "/api/progress")[2])["message"].endswith("207.8 MB")


def test_status_log_for_the_page(server):
    """The exe has no status window of its own; the page shows these lines (2026-11-02 "창 2개를 하나로")."""
    srv, _ = server
    srv.note("Edge 브라우저를 여는 중…")
    srv.note("SharePoint에서 받는 중… Tool list for Field validation 20250325.xlsx")
    srv.note("SharePoint에서 받는 중… 2.6 / 10.7 MB")
    srv.note("SharePoint에서 받음: Tool list for Field validation 20250325.xlsx")
    body = json.loads(_req(srv, "/api/progress")[2])
    lines = [x["m"] for x in body["log"]]
    assert lines == ["Edge 브라우저를 여는 중…", "SharePoint에서 받는 중… 2.6 / 10.7 MB",
                     "SharePoint에서 받음: Tool list for Field validation 20250325.xlsx"]  # progress kept to one line
    assert body["seq"] == 4 and all(len(x["t"]) == 8 for x in body["log"])


def test_open_folders(server, tmp_path):
    srv, _ = server
    assert _req(srv, "/api/open?what=output", method="POST", token=False)[0] == 403
    status, _, body = _req(srv, "/api/open?what=output", method="POST")
    assert status == 200 and srv.opened == [tmp_path / "out"]
    assert _req(srv, "/api/open?what=logs", method="POST")[0] == 200 and srv.opened[-1] == tmp_path
    assert _req(srv, "/api/open?what=C:%5CWindows", method="POST")[0] == 404 and len(srv.opened) == 2


def test_download_upload_and_save(server):
    srv, fake = server
    status, headers, body = _req(srv, "/api/sp/download", method="POST")
    assert status == 200 and body == fake.data and "Tool%20Verification.xlsx" in headers["X-File-Name"]
    status, _, body = _req(srv, "/api/sp/upload", method="POST", body=b"PKnew")
    assert status == 200 and json.loads(body)["sheets"][-1] == "TEST" and fake.uploaded == b"PKnew"
    fake.fail = True
    status, _, body = _req(srv, "/api/sp/upload", method="POST", body=b"PKnew")
    assert status == 409 and json.loads(body)["code"] == "locked"
    status, _, body = _req(srv, "/api/save?name=a.xlsx", method="POST", body=b"PKx")
    saved = json.loads(body)["path"]
    assert status == 200 and open(saved, "rb").read() == b"PKx"


def test_size_limit_leaves_room_for_the_real_workbook():
    """The SharePoint file was 207.8 MB on 2026-11-02; a 200 MiB limit would soon refuse it."""
    from desktop.server import MAX_UPLOAD
    assert MAX_UPLOAD >= 4 * 208_000_000


# ---------------------------------------------------------------- real browser
CHROMIUM = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"


@pytest.fixture
def browser_env(monkeypatch):
    pytest.importorskip("selenium")
    if not os.path.exists(CHROMIUM):
        pytest.skip("Chromium not available")
    # a chromedriver of another version on PATH would win over Selenium Manager
    monkeypatch.setenv("PATH", os.pathsep.join(p for p in os.environ["PATH"].split(os.pathsep) if "node22" not in p))


def test_sharepoint_through_a_real_browser(browser_env, tmp_path):
    from desktop.sharepoint import BrowserSession, SharePoint

    path = build_workbook(tmp_path / "sp.xlsx", with_pictures=False)
    with open(path, "rb") as fh:
        original = fh.read()
    mock = MockSharePoint(original).start()
    session = BrowserSession(tmp_path / "profile", report=lambda m: None, browser="chrome", binary=CHROMIUM, headless=True)
    try:
        session.open_tool("data:text/html,<p>tool</p>")
        sp = SharePoint(session, SharePointFile.parse(mock.page_url), report=lambda m: None, login_timeout=20)

        with pytest.raises(SharePointError) as e:
            sp.upload(b"PK...")  # nothing downloaded yet
        assert e.value.code == "no_base"

        data, info = sp.download()
        assert data == original and info["Name"] == mock.name and sp.etag == mock.etag
        assert session.driver.current_window_handle == session.tool_window  # back on the tool

        # in real use the page sends the workbook with the TEST sheet added
        changed = original
        result = sp.upload(changed)
        assert result["ok"] and mock.uploads[-1] == changed and result["sheets"] == ["대전", "광주"]
        assert sp.etag == mock.etag  # follows the new version

        mock.bump()  # a co-worker saved meanwhile
        with pytest.raises(SharePointError) as e:
            sp.upload(changed)
        assert e.value.code == "changed" and len(mock.uploads) == 1

        sp.download()  # fresh copy → allowed again, unless the file is open for editing
        mock.locked = True
        with pytest.raises(SharePointError) as e:
            sp.upload(changed)
        assert e.value.code == "locked" and "taesu.kim@philips.com" in str(e.value)
    finally:
        session.quit()
        mock.stop()
        shutil.rmtree(tmp_path / "profile", ignore_errors=True)


def test_slow_transfers_outlast_the_webdriver_command_limit(browser_env, tmp_path, monkeypatch):
    """2026-11-02 on the user's PC: a download longer than the 30 s command limit failed (ReadTimeoutError)."""
    from desktop import sharepoint as spmod

    monkeypatch.setattr(spmod, "COMMAND_TIMEOUT", 5)
    path = build_workbook(tmp_path / "slow.xlsx", with_pictures=True)
    with open(path, "rb") as fh:
        original = fh.read()
    mock = MockSharePoint(original).start()
    mock.download_seconds = 12
    mock.upload_seconds = 8
    messages = []
    session = spmod.BrowserSession(tmp_path / "profile", report=messages.append, browser="chrome", binary=CHROMIUM, headless=True)
    try:
        session.open_tool("data:text/html,<p>tool</p>")
        sp = spmod.SharePoint(session, SharePointFile.parse(mock.page_url), report=messages.append, login_timeout=20)
        data, _ = sp.download()
        assert data == original
        assert any("MB" in m for m in messages)  # progress was shown while receiving
        mock.download_seconds = 0
        result = sp.upload(data)
        assert result["ok"] and mock.uploads[-1] == data
    finally:
        session.quit()
        mock.stop()
        shutil.rmtree(tmp_path / "profile", ignore_errors=True)


def test_program_never_opens_the_workbook_itself(browser_env, tmp_path):
    """2026-11-02: the program opened the file link (Excel web, edit mode) to log in, which locked the
    file in the user's own name, and its upload was refused with 423."""
    from desktop.sharepoint import BrowserSession, SharePoint

    path = build_workbook(tmp_path / "own.xlsx", with_pictures=False)
    with open(path, "rb") as fh:
        original = fh.read()
    mock = MockSharePoint(original).start()  # opening Doc.aspx locks the file, like Excel for the web
    session = BrowserSession(tmp_path / "profile", report=lambda m: None, browser="chrome", binary=CHROMIUM, headless=True)
    try:
        session.open_tool("data:text/html,<p>tool</p>")
        sp = SharePoint(session, SharePointFile.parse(mock.page_url), report=lambda m: None, login_timeout=20)
        data, _ = sp.download()
        assert not mock.locked
        assert not any("Doc.aspx" in r for r in mock.requests)
        assert sp.upload(data)["ok"]

        # a SharePoint window left on the workbook (older version) is moved off it before uploading
        session.driver.switch_to.window(session.sp_window)
        session.driver.get(mock.page_url)
        mock.locked = False  # the lock is released once that page is left
        session.show_tool()
        sp.download()
        assert sp.upload(data)["ok"]
    finally:
        session.quit()
        mock.stop()
        shutil.rmtree(tmp_path / "profile", ignore_errors=True)


def test_program_ends_when_the_tool_window_is_closed(browser_env, tmp_path, monkeypatch):
    """One window (2026-11-02): the page shows the status log, and closing it ends the program."""
    import threading
    import time

    from desktop import app as desk

    monkeypatch.setattr(desk, "APP_DIR", tmp_path / "home")
    monkeypatch.setattr(desk, "LOG_PATH", tmp_path / "home" / "debug_log.txt")
    monkeypatch.setattr(desk, "OUTPUT_DIR", tmp_path / "home" / "output")
    path = build_workbook(tmp_path / "w.xlsx", with_pictures=False)
    with open(path, "rb") as fh:
        mock = MockSharePoint(fh.read()).start()
    args = desk.parse_args(["--browser", "chrome", "--binary", CHROMIUM, "--headless", "--no-gui",
                            "--sharepoint-url", mock.page_url, "--profile", str(tmp_path / "profile")])
    program = desk.Program(args)
    # Headless Chromium sometimes takes 15 s to close the very first window of the browser, even an empty one,
    # so the tool page goes into a second window here. (A real Edge window closed with X is not affected.)
    program.session.start()
    program.session.driver.switch_to.new_window("window")
    result = []
    runner = threading.Thread(target=lambda: result.append(program.run()), daemon=True)
    try:
        runner.start()
        session = program.session

        def page_log():
            with session.lock:
                if session.tool_window is None:
                    return ""
                session.driver.switch_to.window(session.tool_window)
                return session.driver.execute_script(
                    "const e = document.getElementById('program-log'); return e && !document.getElementById('program').hidden"
                    " ? e.innerText : ''") or ""

        end = time.time() + 60
        while "SharePoint에서 받음" not in page_log():  # the page downloads by itself and shows the log
            assert time.time() < end, page_log()
            time.sleep(0.5)
        assert "도구 화면을 열었습니다" in page_log()
        assert runner.is_alive()

        with session.lock:  # the user closes the tool window; the SharePoint window is still open
            assert session.sp_window in session.driver.window_handles
            session.driver.switch_to.window(session.tool_window)
            session.driver.close()
        runner.join(20)
        assert not runner.is_alive() and result == [0]
        assert session.driver is None  # Edge (with the SharePoint window) was closed too
    finally:
        program.stop()
        mock.stop()
        shutil.rmtree(tmp_path / "profile", ignore_errors=True)


def test_sharepoint_window_stays_out_of_the_way(browser_env, tmp_path):
    """2026-11-02: the 'Site contents' window was always on screen, and closing it broke the transfer.
    Now it stays minimized, survives being closed, and shows the uploaded file (view mode) at the end."""
    import threading
    import time

    from desktop.sharepoint import BrowserSession, SharePoint

    path = build_workbook(tmp_path / "win.xlsx", with_pictures=False)
    with open(path, "rb") as fh:
        original = fh.read()
    mock = MockSharePoint(original).start()
    messages = []
    session = BrowserSession(tmp_path / "profile", report=messages.append, browser="chrome", binary=CHROMIUM, headless=True)
    try:
        session.start()
        session.driver.switch_to.new_window("window")  # see test_program_ends_when_the_tool_window_is_closed
        session.open_tool("data:text/html,<p>tool</p>")
        sp = SharePoint(session, SharePointFile.parse(mock.page_url), report=messages.append, login_timeout=20)

        def state(handle):
            with session.lock:
                return session.driver.execute_cdp_cmd("Browser.getWindowForTarget", {"targetId": handle})["bounds"]["windowState"]

        data, _ = sp.download()
        assert state(session.sp_window) == "minimized" and state(session.tool_window) != "minimized"
        assert session.driver.current_window_handle == session.tool_window
        with session.lock:
            session.driver.switch_to.window(session.sp_window)
            assert "닫지 마세요" in session.driver.title  # says what it is, if the user opens it from the taskbar
            session.show_tool()

        # the user closes the SharePoint window in the middle of a download: it is opened again and finishes
        mock.download_seconds = 6
        closed = session.sp_window
        threading.Timer(2.5, lambda: session.driver.execute_cdp_cmd("Target.closeTarget", {"targetId": closed})).start()
        data2, _ = sp.download()
        assert data2 == original and session.sp_window != closed
        assert any("다시 여는 중" in m for m in messages)
        mock.download_seconds = 0

        # after an upload the file itself is shown, for viewing only, so it is not locked
        result = sp.upload(data2)
        assert result["ok"] and result["opened"] and mock.views == 1 and not mock.locked
        with session.lock:
            assert session.driver.current_window_handle == session.sp_window
            assert "Doc.aspx?action=view" in session.driver.current_url
        assert state(session.sp_window) != "minimized"

        # the next round leaves the workbook page first, and the upload is not blocked
        data3, _ = sp.download()
        assert state(session.sp_window) == "minimized"
        assert sp.upload(data3)["ok"] and not mock.locked
    finally:
        session.quit()
        mock.stop()
        shutil.rmtree(tmp_path / "profile", ignore_errors=True)
