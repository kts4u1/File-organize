"""A small stand-in for SharePoint Online, for testing desktop/sharepoint.py.

Implements only what the program uses, with SharePoint's rules that matter:
requests need the login cookie the file page sets, uploads need the form
digest from /_api/contextinfo, If-Match must match the current ETag (412),
and a file open for editing is locked (423).
"""

from __future__ import annotations

import json
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

SITE = "/sites/HSCSCCJ-jr-all"
GUID = "5b1d0df6-6a62-49fb-92e3-3cb102684f84"
DIGEST = "0x5EED,02 Nov 2026 10:00:00 -0000"


class MockSharePoint:
    def __init__(self, data: bytes, name: str = "Tool_Verification_2026_Q1.xlsx"):
        self.data = data
        self.name = name
        self.version = 7
        self.locked = False
        self.download_seconds = 0.0  # >0: send the file slowly, like a slow network
        self.upload_seconds = 0.0  # >0: answer the upload late
        self.uploads: list[bytes] = []
        self.requests: list[str] = []
        self.httpd = ThreadingHTTPServer(("127.0.0.1", 0), self._handler())
        self.httpd.daemon_threads = True

    @property
    def origin(self) -> str:
        return f"http://127.0.0.1:{self.httpd.server_address[1]}"

    @property
    def page_url(self) -> str:
        return (f"{self.origin}/:x:/r{SITE}/_layouts/15/Doc.aspx?action=edit&sourcedoc=%7B{GUID}%7D"
                f"&wdExp=TEAMS-TREATMENT&web=1")

    @property
    def etag(self) -> str:
        return f'"{{{GUID.upper()}}},{self.version}"'

    def bump(self, data: bytes | None = None):
        """Someone else saved the file."""
        if data is not None:
            self.data = data
        self.version += 1

    def start(self):
        threading.Thread(target=self.httpd.serve_forever, daemon=True).start()
        return self

    def stop(self):
        self.httpd.shutdown()
        self.httpd.server_close()

    def _handler(self):
        sp = self

        class H(BaseHTTPRequestHandler):
            def log_message(self, *a):
                pass

            def _send(self, status, body=b"", ctype="text/plain", headers=None):
                self.send_response(status)
                self.send_header("Content-Type", ctype)
                self.send_header("Content-Length", str(len(body)))
                for k, v in (headers or {}).items():
                    self.send_header(k, v)
                self.end_headers()
                self.wfile.write(body)

            def _logged_in(self):
                return "FedAuth=ok" in (self.headers.get("Cookie") or "")

            def do_GET(self):
                u = urlparse(self.path)
                sp.requests.append(f"GET {u.path}")
                if u.path.endswith("/_layouts/15/Doc.aspx"):
                    return self._send(200, b"<html><body>Excel Online (mock)</body></html>", "text/html",
                                      {"Set-Cookie": "FedAuth=ok; Path=/; HttpOnly"})
                if not self._logged_in():
                    return self._send(403, b"Access denied")
                if u.path == f"{SITE}/_layouts/15/download.aspx":
                    if parse_qs(u.query).get("UniqueId", [""])[0].lower() != GUID:
                        return self._send(404, b"not found")
                    headers = {"Content-Disposition": f"attachment; filename*=UTF-8''{sp.name}"}
                    if not sp.download_seconds:
                        return self._send(200, sp.data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers)
                    self.send_response(200)
                    self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    self.send_header("Content-Length", str(len(sp.data)))
                    for k, v in headers.items():
                        self.send_header(k, v)
                    self.end_headers()
                    pieces = 24
                    step = -(-len(sp.data) // pieces)
                    for i in range(0, len(sp.data), step):
                        self.wfile.write(sp.data[i:i + step])
                        self.wfile.flush()
                        time.sleep(sp.download_seconds / pieces)
                    return
                if u.path == f"{SITE}/_api/web/GetFileById('{GUID}')":
                    body = {"Name": sp.name, "ServerRelativeUrl": f"{SITE}/Shared Documents/{sp.name}",
                            "ETag": sp.etag, "TimeLastModified": "2026-11-02T01:23:45Z", "Length": str(len(sp.data))}
                    return self._send(200, json.dumps(body).encode(), "application/json")
                return self._send(404, b"not found")

            def do_POST(self):
                u = urlparse(self.path)
                n = int(self.headers.get("Content-Length") or 0)
                body = self.rfile.read(n) if n else b""
                sp.requests.append(f"POST {u.path}")
                if not self._logged_in():
                    return self._send(403, b"Access denied")
                if u.path == f"{SITE}/_api/contextinfo":
                    return self._send(200, json.dumps({"FormDigestValue": DIGEST}).encode(), "application/json")
                if u.path == f"{SITE}/_api/web/GetFileById('{GUID}')/$value":
                    if self.headers.get("X-HTTP-Method") != "PUT":
                        return self._send(400, b"expected X-HTTP-Method: PUT")
                    if self.headers.get("X-RequestDigest") != DIGEST:
                        return self._send(403, b"The security validation for this page is invalid.")
                    if sp.locked:
                        msg = (f'<?xml version="1.0" encoding="utf-8"?><m:error xmlns:m="http://schemas.microsoft.com/ado/2007/08/'
                               f'dataservices/metadata"><m:code>-2147018894, Microsoft.SharePoint.SPFileLockException</m:code>'
                               f'<m:message xml:lang="en-US">The file https://share.philips.com{SITE}/Shared Documents/{sp.name} '
                               f'is locked for shared use by taesu.kim@philips.com [membership].</m:message></m:error>')
                        return self._send(423, msg.encode(), "application/xml")
                    match = self.headers.get("If-Match")
                    if match and match != sp.etag:
                        return self._send(412, b'{"error":{"message":"The version of the file has changed."}}', "application/json")
                    time.sleep(sp.upload_seconds)
                    sp.uploads.append(body)
                    sp.data = body
                    sp.version += 1
                    return self._send(204)
                return self._send(404, b"not found")

        return H
