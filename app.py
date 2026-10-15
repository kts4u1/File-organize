"""툴 검교정 관리 — web UI.

Run:  streamlit run app.py   (or run.bat on Windows)
"""

from __future__ import annotations

import datetime as dt
import hashlib
import os

import pandas as pd
import streamlit as st

from toolcal.config import Settings, load_settings, save_settings
from toolcal.imaging import SUPPORTED_EXTENSIONS
from toolcal.pipeline import (
    KIND_LABELS,
    STATUS_LABELS,
    analyze_photos,
    applicable,
    apply_jobs,
    due_status,
    evaluate_all,
    fmt,
)
from toolcal.recognize import ClaudeRecognizer, NullRecognizer
from toolcal.workbook import ToolWorkbook, WorkbookError

st.set_page_config(page_title="툴 검교정 관리", page_icon="🔧", layout="wide")

STATUS_COLORS = {"ok": "green", "review": "orange", "error": "red", "skip": "gray"}
NO_TARGET = "(대상 선택 안 함)"
KIND_KEYS = list(KIND_LABELS)

ss = st.session_state
ss.setdefault("jobs", [])
ss.setdefault("result", None)
ss.setdefault("uploader_gen", 0)


# ------------------------------------------------------------------ sidebar
def sidebar() -> tuple[Settings, bytes | None, str | None]:
    settings = load_settings()
    with st.sidebar:
        st.header("설정")
        excel_path = st.text_input(
            "엑셀 파일 경로",
            value=settings.excel_path,
            placeholder=r"C:\Users\...\Tool_Verification_2026_Q1.xlsx",
            help="이 파일에 바로 저장합니다 (저장 전 백업). OneDrive/SharePoint 동기화 폴더 경로도 됩니다.",
        )
        uploaded = st.file_uploader(
            "또는 엑셀 파일 올리기",
            type=["xlsx"],
            help="올린 파일로 작업하면 결과는 다운로드 버튼으로 받습니다.",
        )
        api_key = st.text_input(
            "Claude API 키",
            value=settings.anthropic_api_key,
            type="password",
            help="비워두면 ANTHROPIC_API_KEY 환경변수를 씁니다. 키가 없으면 대상 행/날짜를 직접 입력합니다.",
        )
        with st.expander("고급 설정"):
            model = st.text_input("모델", value=settings.model)
            effort = st.selectbox("인식 정밀도(effort)", ["low", "medium", "high"],
                                  index=["low", "medium", "high"].index(settings.effort)
                                  if settings.effort in ("low", "medium", "high") else 1)
            archive = st.checkbox("원본 사진 보관 (지역/날짜_이름_시리얼.jpg)", value=settings.archive_photos)
            archive_dir = st.text_input("사진 보관 폴더", value=settings.archive_dir,
                                        placeholder="비우면 엑셀 폴더\\검교정사진")
            backup = st.checkbox("저장 전 엑셀 백업", value=settings.backup)
            warn_days = st.number_input("임박 기준(일)", min_value=1, max_value=365, value=settings.warn_days)

        settings = Settings(
            excel_path=excel_path.strip().strip('"'),
            anthropic_api_key=api_key.strip(),
            model=model.strip() or settings.model,
            effort=effort,
            archive_photos=archive,
            archive_dir=archive_dir.strip().strip('"'),
            backup=backup,
            backup_dir=settings.backup_dir,
            warn_days=int(warn_days),
        )
        if st.button("설정 저장", width="stretch"):
            save_settings(settings)
            st.success("config.json 에 저장했습니다.")

    if uploaded is not None:
        return settings, uploaded.getvalue(), None
    if settings.excel_path:
        if not os.path.isfile(settings.excel_path):
            st.sidebar.error("엑셀 파일을 찾을 수 없습니다.")
            return settings, None, None
        with open(settings.excel_path, "rb") as fh:
            return settings, fh.read(), settings.excel_path
    return settings, None, None


@st.cache_resource(show_spinner="엑셀 읽는 중…", max_entries=4)
def open_workbook(digest: str, _data: bytes) -> ToolWorkbook:  # digest keys the cache
    return ToolWorkbook(_data)


# ------------------------------------------------------------------ helpers
def status_badge(status: str) -> str:
    return f":{STATUS_COLORS[status]}[**{STATUS_LABELS[status]}**]"


def option_list(job, wb: ToolWorkbook) -> list[str]:
    """Matched candidates first, then every other tool."""
    labels = [c.record.label() for c in job.candidates[:5]]
    labels += [r.label() for r in wb.records if r.label() not in labels]
    return [NO_TARGET] + labels


def sync_job(i: int, job, wb: ToolWorkbook) -> bool:
    """Pull this job's widget values into the job. Returns True if the user changed the row.

    Widget state may only be written before the widget is created in a run,
    so this runs before render_job().
    """
    by_label = {r.label(): r for r in wb.records}
    kind_key, target_key, last_key, next_key, inc_key = (f"{n}_{i}" for n in ("kind", "target", "last", "next", "inc"))

    if kind_key in ss:
        job.kind = ss[kind_key]
    if target_key in ss:
        chosen = by_label.get(ss[target_key])
        current = job.target.label() if job.target else NO_TARGET
        if ss[target_key] != current:
            job.target = chosen
            job.target_confirmed = chosen is not None
            return True
    if last_key in ss:
        job.cal_last = ss[last_key]
    if next_key in ss:
        job.cal_next = ss[next_key]
    if inc_key in ss:
        job.include = ss[inc_key]
    return False


def render_job(i: int, job, wb: ToolWorkbook) -> None:
    rec = job.recognition
    with st.container(border=True):
        head, badge = st.columns([5, 1])
        head.markdown(f"**{i + 1}. {job.filename}**")
        badge.markdown(status_badge(job.status))

        pic, info, edit = st.columns([2, 2, 3])
        pic.image(job.picture, width="stretch",
                  caption="엑셀에 들어갈 사진" + (" (교정필증 자동 크롭)" if job.picture is job.label_crop else ""))

        with info:
            st.caption("사진에서 읽은 값")
            if rec:
                st.markdown(
                    f"- 이름: **{rec.engineer_name or '-'}**\n"
                    f"- 시리얼: **{rec.serial_number or '-'}**"
                    + (f" (다른 판독: {', '.join(rec.serial_alternatives)})" if rec.serial_alternatives else "")
                    + f"\n- 교정일: **{rec.calibration_date or '-'}**\n"
                    f"- 차기교정일: **{rec.next_due_date or '-'}**\n"
                    f"- 성적서번호: {rec.certificate_no or '-'}\n"
                    f"- 확신도: {rec.confidence}"
                )
                if rec.notes:
                    st.caption(rec.notes)

        with edit:
            st.selectbox("사진 종류", KIND_KEYS, key=f"kind_{i}", format_func=lambda k: KIND_LABELS[k])
            options = option_list(job, wb)
            if ss.get(f"target_{i}") not in options:
                ss[f"target_{i}"] = NO_TARGET
            st.selectbox("엑셀 대상 행", options, key=f"target_{i}")
            if job.kind == "tool_photo":
                d1, d2 = st.columns(2)
                d1.date_input("교정일 (Cal Last)", key=f"last_{i}", format="YYYY-MM-DD")
                d2.date_input("차기교정일 (Cal Next)", key=f"next_{i}", format="YYYY-MM-DD")
            st.checkbox("엑셀에 반영", key=f"inc_{i}", disabled=job.status in ("error", "skip"))

        for message in job.messages:
            st.markdown(f"- {message}")


def init_widget_state(jobs) -> None:
    for i, job in enumerate(jobs):
        ss[f"kind_{i}"] = job.kind
        ss[f"target_{i}"] = job.target.label() if job.target else NO_TARGET
        ss[f"last_{i}"] = job.cal_last
        ss[f"next_{i}"] = job.cal_next
        ss[f"inc_{i}"] = job.include


def reset_jobs() -> None:
    for key in list(ss.keys()):
        if key.split("_")[0] in ("kind", "target", "last", "next", "inc"):
            del ss[key]
    ss.jobs = []
    ss.result = None
    ss.uploader_gen += 1


# ------------------------------------------------------------------ pages
def page_update(settings: Settings, data: bytes, path: str | None, wb: ToolWorkbook) -> None:
    if ss.result:
        show_result(ss.result)
        if st.button("새 작업 시작", type="primary"):
            reset_jobs()
            st.rerun()
        return

    files = st.file_uploader(
        "검교정 후 찍은 툴 사진 (이름·시리얼 메모 + 교정필증) 또는 SMAX 캡쳐를 올리세요",
        type=list(SUPPORTED_EXTENSIONS),
        accept_multiple_files=True,
        key=f"photos_{ss.uploader_gen}",
    )
    api_key = settings.api_key()
    if not api_key:
        st.warning("Claude API 키가 없어 자동 인식은 하지 않습니다. 사진마다 대상 행과 날짜를 직접 선택하세요.")

    if files and st.button(f"사진 {len(files)}장 인식 시작", type="primary", disabled=bool(ss.jobs)):
        recognizer = (ClaudeRecognizer(api_key=api_key, model=settings.model, effort=settings.effort)
                      if api_key else NullRecognizer())
        bar = st.progress(0.0, text="인식 준비 중…")

        def progress(done, total, name):
            bar.progress(done / total, text=f"{done}/{total} 완료 — {name}")

        ss.jobs = analyze_photos(wb, [(f.name, f.getvalue()) for f in files], recognizer, progress)
        init_widget_state(ss.jobs)
        bar.empty()

    if not ss.jobs:
        return

    retargeted = [sync_job(i, job, wb) for i, job in enumerate(ss.jobs)]
    evaluate_all(ss.jobs)
    for i, job in enumerate(ss.jobs):
        # Dates may have been derived from the interval; errors cannot be applied.
        ss[f"last_{i}"], ss[f"next_{i}"] = job.cal_last, job.cal_next
        if retargeted[i]:
            # A hand-picked row is ticked only if nothing else still needs a look.
            ss[f"inc_{i}"] = job.include = job.status == "ok"
        if job.status in ("error", "skip"):
            ss[f"inc_{i}"] = job.include = False

    counts = {s: sum(j.status == s for j in ss.jobs) for s in STATUS_LABELS}
    st.markdown(" · ".join(f"{status_badge(s)} {n}건" for s, n in counts.items() if n))
    for i, job in enumerate(ss.jobs):
        render_job(i, job, wb)

    selected = [j for j in ss.jobs if applicable(j)]
    c1, c2 = st.columns([1, 3])
    target_text = f"{os.path.basename(path)} 에 저장" if path else "결과 파일 만들기"
    if c1.button(f"엑셀에 반영 ({len(selected)}건)", type="primary", disabled=not selected):
        apply_to_excel(settings, data, path)
        st.rerun()
    if c2.button("취소하고 다시 올리기"):
        reset_jobs()
        st.rerun()
    c2.caption(target_text)


def apply_to_excel(settings: Settings, data: bytes, path: str | None) -> None:
    # Re-read the file right before writing so edits made meanwhile are kept.
    if path:
        with open(path, "rb") as fh:
            data = fh.read()
    wb = ToolWorkbook(data, path)
    archive = settings.resolved_archive_dir(path) if (path and settings.archive_photos) else None
    if not path and settings.archive_photos and settings.archive_dir:
        archive = settings.archive_dir
    changes = apply_jobs(wb, ss.jobs, archive)
    backup = None
    try:
        if path:
            backup = wb.save(path, backup_dir=settings.resolved_backup_dir(path))
            output = None
        else:
            output = wb.to_bytes()
    except WorkbookError as exc:
        st.error(str(exc))
        return
    ss.result = {"changes": changes, "backup": backup, "path": path, "output": output, "archive": archive,
                 "skipped": [j for j in ss.jobs if not applicable(j)]}
    open_workbook.clear()


def show_result(result: dict) -> None:
    changes = result["changes"]
    st.success(f"{len(changes)}건을 엑셀에 반영했습니다." + (f"  →  {result['path']}" if result["path"] else ""))
    if changes:
        st.dataframe(
            pd.DataFrame([{
                "파일": c.filename,
                "대상": c.record.label(),
                "종류": KIND_LABELS[c.kind],
                "교정일": f"{fmt(c.old_cal_last)} → {fmt(c.new_cal_last)}" if c.kind == "tool_photo" else "",
                "차기교정일": f"{fmt(c.old_cal_next)} → {fmt(c.new_cal_next)}" if c.kind == "tool_photo" else "",
                "사진": f"교체 ({c.pictures_replaced}장 삭제)" if c.pictures_replaced else "추가",
                "비고": " ".join(c.notes),
            } for c in changes]),
            hide_index=True,
            width="stretch",
        )
    if result["output"]:
        st.download_button("수정된 엑셀 다운로드", result["output"], file_name="Tool_Verification_updated.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    if result["backup"]:
        st.caption(f"백업: {result['backup']}")
    if result["archive"]:
        st.caption(f"원본 사진 보관: {result['archive']}")
    if result["skipped"]:
        st.warning("반영하지 않은 사진: " + ", ".join(j.filename for j in result["skipped"]))


def page_status(settings: Settings, wb: ToolWorkbook) -> None:
    rows = due_status(wb.records, warn_days=settings.warn_days)
    df = pd.DataFrame(rows)
    regions = sorted(df["지역"].unique()) if not df.empty else []
    picked = st.multiselect("지역", regions, default=regions)
    df = df[df["지역"].isin(picked)]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("전체 툴", len(df))
    m2.metric("만료", int((df["상태"] == "만료").sum()))
    m3.metric(f"{settings.warn_days}일 이내", int((df["상태"] == "임박").sum()))
    m4.metric("정상", int((df["상태"] == "정상").sum()))

    colors = {"만료": "#fde2e1", "임박": "#fff1cc", "정상": "", "날짜 없음": "#eeeeee", "폐기": "#eeeeee"}
    styled = df.style.apply(lambda row: [f"background-color: {colors[row['상태']]}"] * len(row), axis=1)
    st.dataframe(
        styled,
        hide_index=True,
        width="stretch",
        column_config={
            "차기교정일": st.column_config.DateColumn(format="YYYY-MM-DD"),
            "교정일": st.column_config.DateColumn(format="YYYY-MM-DD"),
            "남은 일수": st.column_config.NumberColumn(format="%d일"),
        },
    )
    st.caption(f"기준일 {dt.date.today().isoformat()} · 엑셀 'Cal Next' 열 기준")


def main() -> None:
    st.title("🔧 툴 검교정 관리")
    settings, data, path = sidebar()
    if data is None:
        st.info("왼쪽에서 Tool Verification 엑셀 파일 경로를 입력하거나 파일을 올려주세요.")
        return
    try:
        wb = open_workbook(hashlib.sha1(data).hexdigest(), data)
    except WorkbookError as exc:
        st.error(str(exc))
        return
    except Exception as exc:  # unreadable / not an xlsx
        st.error(f"엑셀 파일을 읽을 수 없습니다: {exc}")
        return
    st.caption(f"툴 {len(wb.records)}개 · 시트: {', '.join(wb.sheets)}")

    tab_update, tab_status = st.tabs(["📷 사진으로 업데이트", "📅 Due date 현황"])
    with tab_update:
        page_update(settings, data, path, wb)
    with tab_status:
        page_status(settings, wb)


main()
