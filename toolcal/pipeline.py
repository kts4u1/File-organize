"""Photo -> recognition -> matching -> workbook update."""

from __future__ import annotations

import csv
import datetime as dt
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from typing import Callable, Iterable

from PIL import Image

from .imaging import encode_jpeg, encode_png, load_image, rotate_clockwise
from .label_crop import box_from_fractions, crop_label, find_label
from .matcher import Candidate, add_months, decide, find_candidates, normalize_serial, parse_date
from .recognize import Recognition, RecognitionError, Recognizer
from .workbook import ToolRecord, ToolWorkbook

KIND_LABELS = {"tool_photo": "툴 사진", "smax_screenshot": "SMAX 캡쳐", "other": "기타"}
STATUS_LABELS = {"ok": "자동 반영", "review": "확인 필요", "error": "오류", "skip": "제외"}

LABEL_PHOTO_SIDE = 1000  # px, long side of the label picture stored in Excel
SMAX_PHOTO_SIDE = 1600
DATE_TOLERANCE_DAYS = 3


@dataclass
class PhotoJob:
    filename: str
    image: Image.Image
    recognition: Recognition | None = None
    error: str = ""
    kind: str = "other"
    candidates: list[Candidate] = field(default_factory=list)
    target: ToolRecord | None = None
    target_confirmed: bool = False  # the user picked the row by hand
    match_status: str = "error"
    match_messages: list[str] = field(default_factory=list)
    cal_last: dt.date | None = None
    cal_next: dt.date | None = None
    label_crop: Image.Image | None = None
    status: str = "error"
    messages: list[str] = field(default_factory=list)
    include: bool = False

    @property
    def picture(self) -> Image.Image:
        """What goes into Excel: the cropped label for tool photos, the full image otherwise."""
        if self.kind == "tool_photo" and self.label_crop is not None:
            return self.label_crop
        return self.image


def analyze_photo(wb: ToolWorkbook, filename: str, data: bytes, recognizer: Recognizer) -> PhotoJob:
    try:
        image = load_image(data)
    except Exception as exc:
        return PhotoJob(filename, Image.new("RGB", (8, 8)), error=str(exc), status="error",
                        messages=[str(exc)])
    job = PhotoJob(filename, image)
    known_names = sorted({n for r in wb.records for n in r.names})
    try:
        job.recognition = recognizer.recognize(image, known_names)
    except RecognitionError as exc:
        job.error = str(exc)
        job.recognition = Recognition(image_type="tool_photo")
    rec = job.recognition
    job.kind = rec.image_type
    if rec.rotation:
        job.image = rotate_clockwise(image, rec.rotation)

    if job.kind == "tool_photo":
        # label_box refers to the image as sent, so crop first and turn upright afterwards
        hint = box_from_fractions(image, *rec.label_box) if rec.label_box else None
        box = find_label(image, hint)
        if box is not None:
            job.label_crop = rotate_clockwise(crop_label(image, box), rec.rotation)
        job.cal_last = parse_date(rec.calibration_date)
        job.cal_next = parse_date(rec.next_due_date)

    job.candidates = find_candidates(rec, wb.records)
    decision = decide(job.candidates)
    job.target = decision.candidate.record if decision.candidate else None
    job.match_status = decision.status
    job.match_messages = decision.messages
    evaluate(job)
    job.include = job.status == "ok"
    return job


def analyze_photos(
    wb: ToolWorkbook,
    files: Iterable[tuple[str, bytes]],
    recognizer: Recognizer,
    progress: Callable[[int, int, str], None] | None = None,
    max_workers: int = 4,
) -> list[PhotoJob]:
    """Recognise several photos in parallel (the API calls are network-bound)."""
    files = list(files)
    jobs: list[PhotoJob | None] = [None] * len(files)
    with ThreadPoolExecutor(max_workers=max(1, max_workers)) as pool:
        futures = {pool.submit(analyze_photo, wb, name, data, recognizer): i
                   for i, (name, data) in enumerate(files)}
        for done, future in enumerate(as_completed(futures), start=1):
            i = futures[future]
            try:
                jobs[i] = future.result()
            except Exception as exc:  # one bad photo must not sink the batch
                jobs[i] = PhotoJob(files[i][0], Image.new("RGB", (8, 8)), error=f"처리 중 오류: {exc}")
            if progress:
                progress(done, len(files), files[i][0])
    evaluate_all(jobs)
    for job in jobs:
        job.include = job.status == "ok"
    return jobs


def set_target(job: PhotoJob, record: ToolRecord | None) -> None:
    """Called when the user picks the target row by hand."""
    job.target = record
    job.target_confirmed = record is not None
    evaluate(job)


def evaluate_all(jobs: list[PhotoJob], today: dt.date | None = None) -> None:
    for job in jobs:
        evaluate(job, today)
    flag_duplicates(jobs)
    check_smax_due_dates(jobs)


def check_smax_due_dates(jobs: list[PhotoJob]) -> None:
    """Warn when a SMAX screenshot shows a due date other than the one being written.

    The due date to compare with is the new one from a tool photo of the same
    tool in this batch, or else the one already in Excel.
    """
    planned = {j.target.key: j.cal_next for j in jobs
               if j.kind == "tool_photo" and j.target is not None and j.cal_next and j.status != "error"}
    for job in jobs:
        if job.kind != "smax_screenshot" or job.target is None or job.recognition is None:
            continue
        shown = parse_date(job.recognition.next_due_date)
        expected = planned.get(job.target.key, job.target.cal_next)
        if shown and expected and shown != expected:
            job.messages.append(
                f"참고: SMAX의 Calibration Due Date({shown})가 엑셀에 들어갈 차기교정일({expected})과 다릅니다. "
                "SMAX를 먼저 수정한 뒤 다시 캡쳐했는지 확인하세요."
            )


def evaluate(job: PhotoJob, today: dt.date | None = None) -> None:
    """(Re)compute status and messages (and fill a missing date from the interval).

    `job.include` is left to the caller: it starts as status == 'ok' and is
    the user's choice afterwards.
    """
    today = today or dt.date.today()
    msgs: list[str] = []
    status = "ok"

    def need_review(message: str) -> None:
        nonlocal status
        msgs.append(message)
        if status == "ok":
            status = "review"

    if job.error:
        msgs.append(job.error)
        status = "review" if job.target_confirmed else "error"
    if job.kind == "other":
        job.status, job.messages = "skip", msgs + ["툴 사진이나 SMAX 캡쳐가 아니어서 제외했습니다."]
        return

    if job.target_confirmed:
        msgs.append(f"대상(직접 선택): {job.target.label()}")
    else:
        msgs.extend(job.match_messages)
        if job.match_status == "error":
            status = "error"
        elif job.match_status == "review":
            need_review("매칭 결과를 확인하세요.")

    rec = job.target
    if rec is None:
        job.status, job.messages = "error", msgs
        return

    if job.kind == "tool_photo":
        interval = rec.interval_months
        if job.cal_last and not job.cal_next and interval:
            job.cal_next = add_months(job.cal_last, interval)
            msgs.append(f"차기교정일을 읽지 못해 교정주기({rec.interval_raw})로 계산했습니다.")
        elif job.cal_next and not job.cal_last and interval:
            job.cal_last = add_months(job.cal_next, -interval)
            msgs.append(f"교정일을 읽지 못해 교정주기({rec.interval_raw})로 계산했습니다.")
        if not (job.cal_last and job.cal_next):
            need_review("교정일/차기교정일을 입력하세요.")
        else:
            if job.cal_next <= job.cal_last:
                need_review("차기교정일이 교정일보다 빠릅니다.")
            if job.cal_last > today + dt.timedelta(days=DATE_TOLERANCE_DAYS):
                need_review("교정일이 미래 날짜입니다.")
            if interval:
                expected = add_months(job.cal_last, interval)
                if abs((job.cal_next - expected).days) > DATE_TOLERANCE_DAYS:
                    need_review(f"교정주기({rec.interval_raw})와 차기교정일이 맞지 않습니다 (예상 {expected}).")
            if rec.cal_last and job.cal_last < rec.cal_last:
                need_review(f"엑셀의 기존 교정일({rec.cal_last})보다 이전 날짜입니다. 예전 사진인지 확인하세요.")
            if rec.cal_last == job.cal_last and rec.cal_next == job.cal_next:
                msgs.append("날짜는 엑셀과 같습니다 (사진만 교체).")
            else:
                msgs.append(f"날짜 변경: {fmt(rec.cal_last)} → {fmt(job.cal_last)}, "
                            f"차기 {fmt(rec.cal_next)} → {fmt(job.cal_next)}")
        if job.label_crop is None:
            msgs.append("교정필증을 찾지 못해 사진 전체를 넣습니다.")
    elif job.kind == "smax_screenshot":
        msgs.append("SMAX 캡쳐를 'SMAX Tool Photo' 열에 넣습니다.")

    job.status, job.messages = status, msgs


def flag_duplicates(jobs: list[PhotoJob]) -> None:
    seen: dict[tuple[str, str], list[PhotoJob]] = {}
    for job in jobs:
        if job.target is not None and job.kind != "other":
            seen.setdefault((job.target.key, job.kind), []).append(job)
    for group in seen.values():
        if len(group) > 1:
            names = ", ".join(j.filename for j in group)
            for job in group:
                job.messages.append(f"같은 툴에 대한 사진이 여러 장입니다: {names}")
                if job.status == "ok":
                    job.status = "review"


def fmt(value: dt.date | None) -> str:
    return value.isoformat() if value else "-"


# ------------------------------------------------------------------ apply
@dataclass
class Change:
    filename: str
    record: ToolRecord
    kind: str
    old_cal_last: dt.date | None
    old_cal_next: dt.date | None
    new_cal_last: dt.date | None
    new_cal_next: dt.date | None
    pictures_replaced: int
    notes: list[str] = field(default_factory=list)
    archived_to: str = ""


def applicable(job: PhotoJob) -> bool:
    if not job.include or job.target is None or job.status in ("error", "skip"):
        return False
    if job.kind == "tool_photo":
        return bool(job.cal_last and job.cal_next)
    return job.kind == "smax_screenshot"


def apply_jobs(wb: ToolWorkbook, jobs: list[PhotoJob], archive_dir: str | None = None) -> list[Change]:
    """Write every included job into the (in-memory) workbook. Call wb.save() afterwards.

    `wb` may be a fresh copy of the file the jobs were analysed against; rows
    are re-checked by serial number so a row inserted meanwhile cannot shift
    an update onto the wrong tool.
    """
    changes = []
    for job in jobs:
        if not applicable(job):
            continue
        rec = wb.find_record(job.target.sheet, job.target.row)
        if rec is None or normalize_serial(rec.serial) != normalize_serial(job.target.serial):
            job.status = "error"
            job.messages.append("엑셀 파일이 그 사이 바뀌어 대상 행을 확인할 수 없습니다. 다시 인식하세요.")
            continue
        old_last, old_next = rec.cal_last, rec.cal_next
        notes = []
        if job.kind == "tool_photo":
            for key, value in (("cal_last", job.cal_last), ("cal_next", job.cal_next)):
                if wb.write_date(rec, key, value) == "formula":
                    notes.append(f"{key} 열은 수식이라 그대로 두었습니다.")
            picture = job.picture
            data, ext = encode_jpeg(picture, LABEL_PHOTO_SIDE), "jpeg"
            column = "label_photo"
        else:
            picture = job.image
            data, ext = encode_png(picture, SMAX_PHOTO_SIDE), "png"
            column = "smax_photo"
        size = _encoded_size(picture, LABEL_PHOTO_SIDE if ext == "jpeg" else SMAX_PHOTO_SIDE)
        replaced = 0
        if wb.sheets[rec.sheet].columns.get(column):
            replaced = wb.replace_picture(
                rec, column, data, ext, size,
                name=f"{'Label' if column == 'label_photo' else 'SMAX'} {rec.serial}",
                descr=f"{job.filename} ({dt.date.today().isoformat()})",
            )
        else:
            notes.append(f"'{rec.sheet}' 시트에 사진 열이 없어 사진은 넣지 않았습니다.")
        change = Change(job.filename, rec, job.kind, old_last, old_next,
                        rec.cal_last, rec.cal_next, replaced, notes)
        if archive_dir:
            change.archived_to = archive_photo(archive_dir, job, rec)
        changes.append(change)
    if archive_dir and changes:
        write_log(archive_dir, changes)
    return changes


def _encoded_size(img: Image.Image, max_side: int) -> tuple[int, int]:
    w, h = img.size
    scale = min(1.0, max_side / max(w, h))
    return max(1, round(w * scale)), max(1, round(h * scale))


def _safe(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|\s]+', "_", text).strip("_") or "unknown"


def archive_photo(archive_dir: str, job: PhotoJob, rec: ToolRecord) -> str:
    """Keep the original photo, filed by region and named by date/engineer/serial."""
    folder = os.path.join(archive_dir, _safe(rec.sheet))
    os.makedirs(folder, exist_ok=True)
    date = job.cal_last or dt.date.today()
    name = rec.names[0] if rec.names else (job.recognition.engineer_name if job.recognition else "")
    suffix = "" if job.kind == "tool_photo" else "_SMAX"
    stem = f"{date.isoformat()}_{_safe(name)}_{_safe(rec.serial)}{suffix}"
    path = os.path.join(folder, stem + ".jpg")
    n = 2
    while os.path.exists(path):
        path = os.path.join(folder, f"{stem}_{n}.jpg")
        n += 1
    with open(path, "wb") as fh:
        fh.write(encode_jpeg(job.image, 2400, quality=90))
    return path


def write_log(archive_dir: str, changes: list[Change]) -> str:
    path = os.path.join(archive_dir, "update_log.csv")
    new = not os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        if new:
            writer.writerow(["처리시각", "파일", "지역", "행", "Asset #", "Serial", "유형",
                             "기존 교정일", "기존 차기교정일", "새 교정일", "새 차기교정일", "보관 경로"])
        stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for c in changes:
            writer.writerow([stamp, c.filename, c.record.sheet, c.record.row, c.record.asset, c.record.serial,
                             KIND_LABELS.get(c.kind, c.kind), fmt(c.old_cal_last), fmt(c.old_cal_next),
                             fmt(c.new_cal_last), fmt(c.new_cal_next), c.archived_to])
    return path


# ------------------------------------------------------------------ status
def due_status(records: list[ToolRecord], today: dt.date | None = None, warn_days: int = 30) -> list[dict]:
    today = today or dt.date.today()
    rows = []
    for rec in records:
        if rec.disposed:
            state, days = "폐기", None
        elif rec.cal_next is None:
            state, days = "날짜 없음", None
        else:
            days = (rec.cal_next - today).days
            state = "만료" if days < 0 else ("임박" if days <= warn_days else "정상")
        rows.append({
            "상태": state,
            "남은 일수": days,
            "차기교정일": rec.cal_next,
            "교정일": rec.cal_last,
            "지역": rec.sheet,
            "행": rec.row,
            "담당": ", ".join(rec.names) or rec.asset,
            "Asset #": rec.asset,
            "모델": f"{rec.manufacturer} {rec.model}".strip(),
            "Serial": rec.serial,
        })
    order = {"만료": 0, "임박": 1, "정상": 2, "날짜 없음": 3, "폐기": 4}
    rows.sort(key=lambda r: (order[r["상태"]], r["남은 일수"] if r["남은 일수"] is not None else 10**6))
    return rows
