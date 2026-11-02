# -*- coding: utf-8 -*-
"""툴검교정.exe 빌드 (build.bat에서 호출).  python build.py

PyInstaller 이름은 영문(ToolCal)으로 빌드한 뒤 dist\\툴검교정.exe 로 복사합니다.
(배치 파일에 한글을 쓰면 코드페이지 문제로 깨질 수 있어서 파이썬에서 처리)
"""
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
NAME = "ToolCal"
FINAL = "툴검교정.exe" if sys.platform == "win32" else "툴검교정"
SEP = ";" if sys.platform == "win32" else ":"


def main():
    for f in ("desktop_main.py", "desktop/app.py", "web/index.html", "web/tesseract/manifest.json", "desktop/toolcal.ico"):
        if not (ROOT / f).exists():
            print(f"[오류] {f} 파일이 없습니다. build.bat과 같은 폴더에서 실행해 주세요.")
            return 1
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--onefile", "--noconsole",
        f"--icon={ROOT / 'desktop' / 'toolcal.ico'}",
        f"--name={NAME}",
        "--collect-submodules=selenium",  # 없으면 Edge WebDriver 모듈이 빠짐 (FCO·FSN 프로그램에서 확인)
        f"--add-data={ROOT / 'web'}{SEP}web",
        f"--add-data={ROOT / 'desktop' / 'toolcal.ico'}{SEP}desktop",
        str(ROOT / "desktop_main.py"),
    ]
    print(" ".join(cmd))
    r = subprocess.run(cmd, cwd=ROOT)
    if r.returncode != 0:
        print("[오류] PyInstaller 빌드 실패")
        return r.returncode
    exe = ROOT / "dist" / (NAME + (".exe" if sys.platform == "win32" else ""))
    if not exe.exists():
        print(f"[오류] {exe} 가 만들어지지 않았습니다.")
        return 1
    final = ROOT / "dist" / FINAL
    shutil.copy2(exe, final)
    print(f"\n완료: {final}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
