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


def test_parse_rejects_other_links():
    with pytest.raises(SharePointError):
        SharePointFile.parse("https://share.philips.com/sites/HSCSCCJ-jr-all/Shared%20Documents/x.xlsx")


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
    srv = ToolServer(WEB, tmp_path / "out", fake).start()
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
        assert e.value.code == "locked"
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
