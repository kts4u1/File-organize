"""Command line: batch-process a folder of photos, or list due dates.

    python -m toolcal update --excel Tool_Verification.xlsx photos/
    python -m toolcal update --excel Tool_Verification.xlsx a.jpg b.jpg --dry-run
    python -m toolcal status --excel Tool_Verification.xlsx --days 30
"""

from __future__ import annotations

import argparse
import os
import sys

from .config import load_settings
from .imaging import SUPPORTED_EXTENSIONS
from .pipeline import KIND_LABELS, STATUS_LABELS, analyze_photos, applicable, apply_jobs, due_status, fmt
from .recognize import ClaudeRecognizer
from .workbook import ToolWorkbook, WorkbookError


def _collect(paths: list[str]) -> list[str]:
    files = []
    for path in paths:
        if os.path.isdir(path):
            for name in sorted(os.listdir(path)):
                if name.lower().rsplit(".", 1)[-1] in SUPPORTED_EXTENSIONS:
                    files.append(os.path.join(path, name))
        elif os.path.isfile(path):
            files.append(path)
        else:
            print(f"! 파일을 찾을 수 없습니다: {path}", file=sys.stderr)
    return files


def cmd_update(args, settings) -> int:
    excel = args.excel or settings.excel_path
    if not excel:
        print("엑셀 파일을 --excel 로 지정하세요.", file=sys.stderr)
        return 2
    files = _collect(args.photos)
    if not files:
        print("처리할 사진이 없습니다.", file=sys.stderr)
        return 2
    api_key = settings.api_key()
    if not api_key:
        print("Claude API 키가 없습니다. ANTHROPIC_API_KEY 환경변수나 config.json 에 설정하세요.", file=sys.stderr)
        return 2

    wb = ToolWorkbook.open(excel)
    recognizer = ClaudeRecognizer(api_key=api_key, model=args.model or settings.model, effort=settings.effort)

    def progress(done, total, name):
        print(f"[{done}/{total}] 인식 완료: {os.path.basename(name)}", flush=True)

    inputs = []
    for path in files:
        with open(path, "rb") as fh:
            inputs.append((os.path.basename(path), fh.read()))
    jobs = analyze_photos(wb, inputs, recognizer, progress)

    print()
    for job in jobs:
        if args.include_review and job.status == "review" and job.target is not None:
            job.include = True
        rec = job.recognition
        print(f"■ {job.filename}  [{KIND_LABELS.get(job.kind, job.kind)}] → {STATUS_LABELS[job.status]}"
              f"{'' if applicable(job) else ' (반영 안 함)'}")
        if rec:
            print(f"   인식: 이름={rec.engineer_name or '-'} 시리얼={rec.serial_number or '-'} "
                  f"교정일={rec.calibration_date or '-'} 차기={rec.next_due_date or '-'}")
        for message in job.messages:
            print(f"   - {message}")

    selected = [j for j in jobs if applicable(j)]
    if args.dry_run:
        print(f"\n(미리보기) 반영 대상 {len(selected)}건 — 파일은 수정하지 않았습니다.")
        return 0
    if not selected:
        print("\n반영할 항목이 없습니다.")
        return 1

    out = args.out or excel
    archive = None if args.no_archive else settings.resolved_archive_dir(excel)
    changes = apply_jobs(wb, jobs, archive)
    try:
        backup = wb.save(out, backup_dir=settings.resolved_backup_dir(excel) if out == excel else None)
    except WorkbookError as exc:
        print(f"\n저장 실패: {exc}", file=sys.stderr)
        return 1
    print(f"\n{len(changes)}건 반영 → {out}")
    for c in changes:
        dates = (f"교정일 {fmt(c.old_cal_last)}→{fmt(c.new_cal_last)}, 차기 {fmt(c.old_cal_next)}→{fmt(c.new_cal_next)}"
                 if c.kind == "tool_photo" else "SMAX 캡쳐 교체")
        print(f"   ✓ {c.record.label()}: {dates}")
    if backup:
        print(f"백업: {backup}")
    if archive:
        print(f"사진 보관: {archive}")
    skipped = [j for j in jobs if not applicable(j)]
    if skipped:
        print(f"반영하지 않은 사진 {len(skipped)}건은 위 메시지를 확인하세요 (--include-review 로 '확인 필요' 항목 포함).")
    return 0


def cmd_status(args, settings) -> int:
    excel = args.excel or settings.excel_path
    if not excel:
        print("엑셀 파일을 --excel 로 지정하세요.", file=sys.stderr)
        return 2
    wb = ToolWorkbook.open(excel)
    rows = due_status(wb.records, warn_days=args.days)
    for row in rows:
        if args.all or row["상태"] in ("만료", "임박"):
            days = "" if row["남은 일수"] is None else f"{row['남은 일수']:>5}일"
            print(f"{row['상태']:<4} {days:>7}  {fmt(row['차기교정일'])}  {row['지역']} {row['행']:>3}행  "
                  f"{row['담당']:<8} {row['모델']:<22} S/N {row['Serial']}")
    expired = sum(r["상태"] == "만료" for r in rows)
    soon = sum(r["상태"] == "임박" for r in rows)
    print(f"\n전체 {len(rows)}건 · 만료 {expired}건 · {args.days}일 이내 {soon}건")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="toolcal", description="툴 검교정 사진 → 엑셀 자동 반영")
    sub = parser.add_subparsers(dest="command", required=True)

    p_update = sub.add_parser("update", help="사진을 인식해 엑셀에 반영")
    p_update.add_argument("photos", nargs="+", help="사진 파일 또는 폴더")
    p_update.add_argument("--excel", help="Tool Verification 엑셀 파일")
    p_update.add_argument("--out", help="다른 파일로 저장 (기본: 원본에 저장, 백업 생성)")
    p_update.add_argument("--dry-run", action="store_true", help="인식 결과만 보고 저장하지 않음")
    p_update.add_argument("--include-review", action="store_true", help="'확인 필요' 항목도 반영")
    p_update.add_argument("--no-archive", action="store_true", help="원본 사진 보관 안 함")
    p_update.add_argument("--model", help="Claude 모델 ID")
    p_update.set_defaults(func=cmd_update)

    p_status = sub.add_parser("status", help="Due date 현황")
    p_status.add_argument("--excel", help="Tool Verification 엑셀 파일")
    p_status.add_argument("--days", type=int, default=30, help="임박 기준 일수 (기본 30)")
    p_status.add_argument("--all", action="store_true", help="정상 항목도 모두 표시")
    p_status.set_defaults(func=cmd_status)

    args = parser.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):  # Korean output on a cp949 Windows console
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    try:
        return args.func(args, load_settings())
    except WorkbookError as exc:
        print(f"오류: {exc}", file=sys.stderr)
        return 1
