"""Local web server for the desktop program.

Serves the browser tool (web/) with its libraries from disk, and gives the
page three endpoints the browser cannot do on its own:

  GET  /api/status       → whether SharePoint is configured
  GET  /api/progress     → the latest status line and the recent ones (the program's status log)
  POST /api/sp/download  → the latest Excel from SharePoint (via Edge login)
  POST /api/sp/upload    → put the updated Excel back on SharePoint (?done=ids: then move those phone photos)
  GET  /api/sp/photos    → new phone photos in the SharePoint photo folder (created when missing)
  GET  /api/sp/photo?id= → one of those photos
  POST /api/save         → keep a copy in ~/ToolCal/output
  POST /api/open?what=   → open the output or log folder in Explorer

Only 127.0.0.1 is bound, and every /api call must carry the per-run token
that is written into the served page, so other websites cannot use it.
"""

from __future__ import annotations

import collections
import json
import mimetypes
import secrets
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, unquote, urlparse, parse_qs

from .sharepoint import SharePoint, SharePointError

# The page loads these from CDNs; the desktop build serves bundled copies.
CDN_TO_LOCAL = {
    "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js": "/vendor/jszip.min.js",
    "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js": "/vendor/jspdf.umd.min.js",
    "https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.js": "/vendor/jsQR.js",
    "https://cdn.jsdelivr.net/npm/qrcode-generator@1.4.4/qrcode.js": "/vendor/qrcode.js",
    "https://cdn.jsdelivr.net/npm/tesseract.js@6.0.1/dist/tesseract.min.js": "/vendor/tesseract.min.js",
}
# Same document skeleton the claude.ai artifact host wraps the page in.
SKELETON_HEAD = (
    '<!doctype html><html lang="ko"><head><meta charset="utf-8">'
    '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
    '<meta name="tc-token" content="{token}"><meta name="tc-desktop" content="1">'
    "<style>:root{{color-scheme:light}}body{{margin:0;font:14px system-ui}}img{{max-width:100%}}"
    "[hidden]{{display:none!important}}</style></head><body>"
)
TYPES = {".js": "application/javascript; charset=utf-8", ".json": "application/json", ".html": "text/html; charset=utf-8",
         ".css": "text/css; charset=utf-8", ".png": "image/png", ".svg": "image/svg+xml", ".ico": "image/x-icon"}
# The real workbook is already ~208 MB (2026-11-02) and grows with every photo.
MAX_UPLOAD = 1024 * 1024 * 1024


class ToolServer:
    def __init__(self, web_dir: Path, output_dir: Path, sharepoint: SharePoint | None,
                 report=print, log=lambda m: None, port: int = 0, folders: dict | None = None, opener=None):
        self.web_dir = Path(web_dir).resolve()
        self.output_dir = Path(output_dir)
        self.sharepoint = sharepoint
        self.report = report
        self.log = log
        self.token = secrets.token_urlsafe(24)
        self.last_message = ""
        self.lines: collections.deque = collections.deque(maxlen=100)  # status log shown on the page
        self.seq = 0
        self.folders = {k: Path(v) for k, v in (folders or {}).items()}
        self.opener = opener
        self.httpd = ThreadingHTTPServer(("127.0.0.1", port), self._handler())
        self.httpd.daemon_threads = True
        self.thread: threading.Thread | None = None

    @property
    def url(self) -> str:
        return f"http://127.0.0.1:{self.httpd.server_address[1]}/"

    def start(self):
        self.thread = threading.Thread(target=self.httpd.serve_forever, name="tool-server", daemon=True)
        self.thread.start()
        return self

    def stop(self):
        self.httpd.shutdown()
        self.httpd.server_close()

    def note(self, msg: str):
        """Remember a status line for the page (GET /api/progress).

        Progress updates of the same step ("SharePoint에서 받는 중… 2.6 / 10.7 MB") replace each other
        in the log, so a long transfer stays one line.
        """
        self.last_message = msg
        key = msg.split("…", 1)[0] if "…" in msg else None
        line = {"t": time.strftime("%H:%M:%S"), "m": msg, "k": key}
        if key and self.lines and self.lines[-1]["k"] == key:
            self.lines[-1] = line
        else:
            self.lines.append(line)
        self.seq += 1

    # ------------------------------------------------------------ pages
    def index_html(self) -> bytes:
        page = (self.web_dir / "index.html").read_text(encoding="utf-8")
        for cdn, local in CDN_TO_LOCAL.items():
            page = page.replace(cdn, local)
        return (SKELETON_HEAD.format(token=self.token) + page + "</body></html>").encode("utf-8")

    def static_path(self, path: str) -> Path | None:
        if path == "/tesseract/worker-wrap.js":
            path = "/js/ocr-worker.js"  # published under this name next to the Tesseract worker
        target = (self.web_dir / unquote(path).lstrip("/")).resolve()
        if self.web_dir not in target.parents or not target.is_file():
            return None
        return target

    def save_copy(self, name: str, data: bytes) -> Path:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        safe = "".join(ch for ch in Path(name).name if ch not in '\\/:*?"<>|').strip() or "Tool_Verification.xlsx"
        target = self.output_dir / safe
        if target.exists():
            target = target.with_name(f"{target.stem}_{time.strftime('%Y%m%d_%H%M%S')}{target.suffix}")
        target.write_bytes(data)
        return target

    # ------------------------------------------------------------ handler
    def _handler(self):
        server = self

        class Handler(BaseHTTPRequestHandler):
            protocol_version = "HTTP/1.1"

            def log_message(self, fmt, *args):  # quiet; errors go to the app log
                pass

            def _send(self, status: int, body: bytes, ctype: str, extra: dict | None = None):
                self.send_response(status)
                self.send_header("Content-Type", ctype)
                self.send_header("Content-Length", str(len(body)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("X-Content-Type-Options", "nosniff")
                for k, v in (extra or {}).items():
                    self.send_header(k, v)
                self.end_headers()
                self.wfile.write(body)

            def _json(self, status: int, obj: dict):
                self._send(status, json.dumps(obj, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

            def _authorized(self) -> bool:
                if secrets.compare_digest(self.headers.get("X-TC-Token", ""), server.token):
                    return True
                self._json(403, {"ok": False, "message": "허용되지 않은 요청입니다."})
                return False

            def do_GET(self):
                path = urlparse(self.path).path
                if path in ("/", "/index.html"):
                    return self._send(200, server.index_html(), "text/html; charset=utf-8")
                if path == "/api/status":
                    if not self._authorized():
                        return
                    sp = server.sharepoint
                    return self._json(200, {"ok": True, "desktop": True, "sharepoint": bool(sp),
                                            "name": sp.file.name if sp else "", "url": sp.file.url if sp else "",
                                            "output": str(server.output_dir)})
                if path == "/api/progress":
                    if not self._authorized():
                        return
                    return self._json(200, {"ok": True, "message": server.last_message, "seq": server.seq,
                                            "log": [{"t": x["t"], "m": x["m"]} for x in list(server.lines)]})
                if path in ("/api/sp/photos", "/api/sp/photo"):
                    if not self._authorized():
                        return
                    return self._photos(path)
                target = server.static_path(path)
                if target is None:
                    return self._send(404, b"not found", "text/plain")
                ctype = TYPES.get(target.suffix.lower()) or mimetypes.guess_type(target.name)[0] or "application/octet-stream"
                return self._send(200, target.read_bytes(), ctype)

            def _photos(self, path: str):
                sp = server.sharepoint
                if not sp:
                    return self._json(400, {"ok": False, "message": "SharePoint 주소가 설정되지 않았습니다."})
                try:
                    if path == "/api/sp/photos":
                        return self._json(200, sp.list_photos())
                    uid = (parse_qs(urlparse(self.path).query).get("id") or [""])[0]
                    data, meta = sp.fetch_photo(uid)
                    name = str(meta.get("name") or "photo.jpg")
                    ctype = mimetypes.guess_type(name)[0] or "application/octet-stream"
                    return self._send(200, data, ctype, {"X-File-Name": quote(name)})
                except SharePointError as exc:
                    server.log(f"{path}: {exc} ({exc.code}, {exc.status})")
                    return self._json(502, {"ok": False, "code": exc.code, "status": exc.status, "message": str(exc)})
                except Exception as exc:
                    server.log(f"{path}: 예외 {type(exc).__name__}: {exc}")
                    return self._json(500, {"ok": False, "code": "error", "message": f"처리 중 오류: {exc}"})

            def _body(self) -> bytes:
                n = int(self.headers.get("Content-Length") or 0)
                if n > MAX_UPLOAD:
                    raise SharePointError("파일이 너무 큽니다.")
                return self.rfile.read(n) if n else b""

            def do_POST(self):
                url = urlparse(self.path)
                if not self._authorized():
                    return
                try:
                    if url.path == "/api/sp/download":
                        if not server.sharepoint:
                            return self._json(400, {"ok": False, "message": "SharePoint 주소가 설정되지 않았습니다."})
                        data, info = server.sharepoint.download()
                        name = info.get("Name") or server.sharepoint.file.name or "Tool_Verification.xlsx"
                        server.report(f"SharePoint에서 받음: {name}")
                        return self._send(200, data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                          {"X-File-Name": quote(name), "X-Modified": quote(str(info.get("TimeLastModified", "")))})
                    if url.path == "/api/sp/upload":
                        if not server.sharepoint:
                            return self._json(400, {"ok": False, "message": "SharePoint 주소가 설정되지 않았습니다."})
                        data = self._body()
                        done = [i for i in (parse_qs(url.query).get("done") or [""])[0].split(",") if i.strip()]
                        result = server.sharepoint.upload(data, done) if done else server.sharepoint.upload(data)
                        server.report(f"SharePoint에 올림: {result.get('name')}")
                        return self._json(200, result)
                    if url.path == "/api/open":
                        what = (parse_qs(url.query).get("what") or [""])[0]
                        folder = server.folders.get(what)
                        if folder is None or server.opener is None:
                            return self._json(404, {"ok": False, "message": "없는 폴더입니다."})
                        server.opener(folder)
                        return self._json(200, {"ok": True, "path": str(folder)})
                    if url.path == "/api/save":
                        name = (parse_qs(url.query).get("name") or ["Tool_Verification.xlsx"])[0]
                        target = server.save_copy(name, self._body())
                        return self._json(200, {"ok": True, "path": str(target)})
                    return self._json(404, {"ok": False, "message": "없는 기능입니다."})
                except SharePointError as exc:
                    server.log(f"{url.path}: {exc} ({exc.code}, {exc.status})")
                    return self._json(409 if exc.code in ("changed", "locked", "no_base") else 502,
                                      {"ok": False, "code": exc.code, "status": exc.status, "message": str(exc)})
                except Exception as exc:  # report, never leave the page hanging
                    server.log(f"{url.path}: 예외 {type(exc).__name__}: {exc}")
                    return self._json(500, {"ok": False, "code": "error", "message": f"처리 중 오류: {exc}"})

        return Handler
