"""툴검교정.exe — the photo tool with SharePoint download/upload on the user's PC.

Starts a local server for the browser tool, opens it in Edge (company SSO
profile, like the FSN dashboard program), and keeps a small status window;
closing that window ends the program.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import queue
import subprocess
import sys
import threading
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
        self.messages: "queue.Queue[str]" = queue.Queue()
        cfg = load_config()
        url = args.sharepoint_url or cfg.get("sharepoint_url") or DEFAULT_SHAREPOINT_URL
        self.session = BrowserSession(PROFILE_DIR if not args.profile else Path(args.profile), self.report,
                                      browser=args.browser, binary=args.binary, headless=args.headless, log=log)
        try:
            self.sharepoint = SharePoint(self.session, SharePointFile.parse(url), self.report, log)
        except SharePointError as exc:
            log(f"SharePoint 주소 오류: {exc}")
            self.report(f"SharePoint 주소 오류: {exc}")
            self.sharepoint = None
        self.server = ToolServer(resource_path("web"), OUTPUT_DIR, self.sharepoint, self.report, log, port=args.port)

    def report(self, msg: str):
        log(msg)
        self.messages.put(msg)

    def start(self):
        self.server.start()
        log(f"서버 시작 {self.server.url}")
        self.open_tool()

    def open_tool(self):
        def run():
            try:
                if not self.session.alive():
                    self.session.quit()
                    self.session.start()
                self.session.open_tool(self.server.url)
                self.report("도구 화면을 열었습니다. Edge 창에서 사용하세요.")
            except Exception as exc:
                log("Edge 열기 실패\n" + traceback.format_exc())
                self.report(f"Edge를 열지 못했습니다: {exc}")
        threading.Thread(target=run, daemon=True).start()

    def stop(self):
        self.session.quit()
        try:
            self.server.stop()
        except Exception:
            pass

    # ------------------------------------------------------------ window
    def run_gui(self):
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.title("툴 검교정")
        root.geometry("520x300")
        try:
            root.iconbitmap(str(resource_path("desktop", "toolcal.ico")))
        except Exception:
            pass
        tk.Label(root, text="툴 검교정 정리", font=("Malgun Gothic", 14, "bold")).pack(anchor="w", padx=14, pady=(12, 0))
        tk.Label(root, text="Edge 창에서 사진을 반영하고 SharePoint에 올립니다. 이 창을 닫으면 종료됩니다.",
                 font=("Malgun Gothic", 9), fg="#555").pack(anchor="w", padx=14)
        box = tk.Text(root, height=8, font=("Malgun Gothic", 9), relief="flat", bg="#f4f5f3", wrap="word")
        box.pack(fill="both", expand=True, padx=14, pady=8)
        box.configure(state="disabled")
        bar = tk.Frame(root)
        bar.pack(fill="x", padx=14, pady=(0, 12))
        for text, cmd in (("도구 화면 열기", self.open_tool), ("저장 폴더", lambda: open_folder(OUTPUT_DIR)),
                          ("로그 폴더", lambda: open_folder(APP_DIR)), ("종료", root.destroy)):
            tk.Button(bar, text=text, command=cmd, font=("Malgun Gothic", 9)).pack(side="left", padx=(0, 6))

        def pump():
            while True:
                try:
                    msg = self.messages.get_nowait()
                except queue.Empty:
                    break
                box.configure(state="normal")
                box.insert("end", f"{dt.datetime.now():%H:%M:%S}  {msg}\n")
                box.see("end")
                box.configure(state="disabled")
            root.after(300, pump)

        try:
            self.start()
        except Exception as exc:
            log("시작 실패\n" + traceback.format_exc())
            messagebox.showerror("툴 검교정", f"시작하지 못했습니다:\n{exc}\n\n로그: {LOG_PATH}")
            return 1
        pump()
        root.mainloop()
        self.stop()
        return 0


def parse_args(argv=None):
    p = argparse.ArgumentParser(prog="툴검교정")
    p.add_argument("--sharepoint-url", help="Tool Verification 엑셀의 SharePoint 주소")
    p.add_argument("--browser", choices=["edge", "chrome"], default="edge")
    p.add_argument("--binary", help="브라우저 실행 파일 경로 (테스트용)")
    p.add_argument("--profile", help="브라우저 프로필 폴더 (테스트용)")
    p.add_argument("--headless", action="store_true", help="화면 없이 실행 (테스트용)")
    p.add_argument("--port", type=int, default=0)
    p.add_argument("--no-gui", action="store_true", help="상태 창 없이 서버와 브라우저만 (테스트용)")
    return p.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    APP_DIR.mkdir(parents=True, exist_ok=True)
    log("=== 시작 ===")
    try:
        program = Program(args)
        if args.no_gui:
            program.start()
            print(program.server.url, flush=True)
            threading.Event().wait()
            return 0
        return program.run_gui()
    except Exception as exc:
        log("치명적 오류\n" + traceback.format_exc())
        if not args.no_gui:
            try:
                from tkinter import messagebox
                messagebox.showerror("툴 검교정", f"오류가 났습니다:\n{exc}\n\n로그: {LOG_PATH}")
            except Exception:
                pass
        return 1
