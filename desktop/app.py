"""툴검교정.exe — the photo tool with SharePoint download/upload on the user's PC.

Starts a local server for the browser tool and opens it in Edge (company SSO
profile, like the FSN dashboard program). There is no separate status window:
the page shows the program's status, and closing the Edge window ends it.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

from .server import ToolServer
from .sharepoint import BrowserSession, SharePoint, SharePointError, SharePointFile

APP_DIR = Path.home() / "ToolCal"
LOG_PATH = APP_DIR / "debug_log.txt"
CONFIG_PATH = APP_DIR / "config.json"
OUTPUT_DIR = APP_DIR / "output"
PROFILE_DIR = APP_DIR / "browser_profile"

# The Tool Verification workbook (2026-11-02 사용자 제공). config.json "sharepoint_url"로 바꿀 수 있습니다.
DEFAULT_SHAREPOINT_URL = (
    "https://share.philips.com/:x:/r/sites/HSCSCCJ-jr-all/_layouts/15/Doc.aspx?action=edit"
    "&sourcedoc=%7B5b1d0df6-6a62-49fb-92e3-3cb102684f84%7D&wdExp=TEAMS-TREATMENT&web=1"
)


def log(msg: str):
    try:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        if LOG_PATH.exists() and LOG_PATH.stat().st_size > 2_000_000:
            LOG_PATH.replace(LOG_PATH.with_suffix(".old.txt"))
        with open(LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(f"[{dt.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")
    except Exception:
        pass


def resource_path(*parts: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base.joinpath(*parts)


def load_config() -> dict:
    try:
        if CONFIG_PATH.exists():
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
            if isinstance(data, dict):
                return data
    except Exception:
        log("config.json 읽기 실패\n" + traceback.format_exc())
    return {}


def open_folder(path: Path):
    path.mkdir(parents=True, exist_ok=True)
    if sys.platform == "win32":
        os.startfile(path)  # noqa: S606 - opening a folder for the user
    else:
        subprocess.Popen(["xdg-open", str(path)])


class Program:
    def __init__(self, args):
        self.args = args
        cfg = load_config()
        url = args.sharepoint_url or cfg.get("sharepoint_url") or DEFAULT_SHAREPOINT_URL
        self.session = BrowserSession(PROFILE_DIR if not args.profile else Path(args.profile), self.report,
                                      browser=args.browser, binary=args.binary, headless=args.headless, log=log)
        try:
            self.sharepoint = SharePoint(self.session, SharePointFile.parse(url), self.report, log)
        except SharePointError as exc:
            log(f"SharePoint 주소 오류: {exc}")
            self.sharepoint = None
        self.server = ToolServer(resource_path("web"), OUTPUT_DIR, self.sharepoint, self.report, log, port=args.port,
                                 folders={"output": OUTPUT_DIR, "logs": APP_DIR}, opener=open_folder)
        self.error: Exception | None = None

    def report(self, msg: str):
        """Status line: written to the log and shown in the tool page (no separate window)."""
        log(msg)
        server = getattr(self, "server", None)
        if server is not None:
            server.note(msg)

    def start(self, background: bool = True) -> bool:
        self.server.start()
        log(f"서버 시작 {self.server.url}")
        if self.sharepoint is None:
            self.report("SharePoint 주소가 잘못되어 받기·올리기를 쓸 수 없습니다. config.json의 sharepoint_url을 확인하세요.")
        if background:
            threading.Thread(target=self.open_tool, daemon=True).start()
            return True
        return self.open_tool()

    def open_tool(self) -> bool:
        try:
            if not self.session.alive():
                self.session.quit()
                self.session.start()
            self.session.open_tool(self.server.url)
            self.report("도구 화면을 열었습니다. 이 창을 닫으면 프로그램이 끝납니다.")
            return True
        except Exception as exc:
            log("Edge 열기 실패\n" + traceback.format_exc())
            self.report(f"Edge를 열지 못했습니다: {exc}")
            self.error = exc
            return False

    def wait_until_closed(self, interval: float = 2.0):
        """Return once the tool window is closed (or Edge is gone): that ends the program."""
        misses = 0
        while misses < 2:  # two checks in a row, so one odd answer does not end the program
            time.sleep(interval)
            misses = 0 if self.session.tool_open() else misses + 1
        log("도구 창이 닫혀 종료합니다.")

    def stop(self):
        self.session.quit()
        try:
            self.server.stop()
        except Exception:
            pass

    def run(self) -> int:
        """One window: the tool page in Edge shows the status; closing it ends the program."""
        if not self.start(background=False):
            self.stop()
            show_error(f"Edge를 열지 못했습니다:\n{self.error}\n\n로그: {LOG_PATH}", self.args)
            return 1
        if self.args.no_gui:
            print(self.server.url, flush=True)
        self.wait_until_closed()
        self.stop()
        return 0


def show_error(text: str, args=None):
    if args is not None and args.no_gui:
        print(text, file=sys.stderr)
        return
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("툴 검교정", text, parent=root)
        root.destroy()
    except Exception:
        pass


def parse_args(argv=None):
    p = argparse.ArgumentParser(prog="툴검교정")
    p.add_argument("--sharepoint-url", help="Tool Verification 엑셀의 SharePoint 주소")
    p.add_argument("--browser", choices=["edge", "chrome"], default="edge")
    p.add_argument("--binary", help="브라우저 실행 파일 경로 (테스트용)")
    p.add_argument("--profile", help="브라우저 프로필 폴더 (테스트용)")
    p.add_argument("--headless", action="store_true", help="화면 없이 실행 (테스트용)")
    p.add_argument("--port", type=int, default=0)
    p.add_argument("--no-gui", action="store_true", help="오류 창 대신 콘솔에 출력하고 주소를 출력 (테스트용)")
    return p.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    APP_DIR.mkdir(parents=True, exist_ok=True)
    log("=== 시작 ===")
    try:
        return Program(args).run()
    except Exception as exc:
        log("치명적 오류\n" + traceback.format_exc())
        show_error(f"오류가 났습니다:\n{exc}\n\n로그: {LOG_PATH}", args)
        return 1
