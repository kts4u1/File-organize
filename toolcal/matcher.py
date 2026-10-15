"""Match a recognised photo to a row of the tool list.

The serial number is the key; the handwritten engineer name confirms it.
Because the set of valid serials is known from the workbook, a slightly
misread serial (one wrong digit, O/0, S/5 ...) can still be resolved when
the name agrees and no other tool is equally close.
"""

from __future__ import annotations

import calendar
import datetime as dt
import re
from dataclasses import dataclass, field

from .recognize import Recognition
from .workbook import ToolRecord

_CONFUSABLE = str.maketrans({"O": "0", "Q": "0", "D": "0", "I": "1", "L": "1", "S": "5",
                             "B": "8", "Z": "2", "G": "6"})


def normalize_serial(value: str) -> str:
    return re.sub(r"[^0-9A-Z]", "", str(value or "").upper())


def _loose(value: str) -> str:
    return normalize_serial(value).translate(_CONFUSABLE).lstrip("0")


def _levenshtein(a: str, b: str) -> int:
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, start=1):
        cur = [i]
        for j, cb in enumerate(b, start=1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def serial_distance(reading: str, serial: str) -> float:
    """0 = same serial, 0.5 = same up to look-alike characters/leading zeros, n = edits."""
    a, b = normalize_serial(reading), normalize_serial(serial)
    if not a or not b:
        return 99
    if a == b:
        return 0
    la, lb = _loose(a), _loose(b)
    if la == lb:
        return 0.5
    return _levenshtein(la, lb)


def name_relation(name: str, rec: ToolRecord) -> str:
    """'exact' | 'partial' | 'mismatch' | 'unknown'."""
    name = re.sub(r"\s+", "", name or "")
    if not re.search(r"[가-힣]", name) or not rec.names:
        return "unknown"
    for known in rec.names:
        if name == known or name in known or known in name:
            return "exact"
    for known in rec.names:
        if len(known) == len(name) and sum(x == y for x, y in zip(known, name)) >= len(name) - 1:
            return "partial"
    return "mismatch"


@dataclass
class Candidate:
    record: ToolRecord
    score: float
    distance: float  # serial distance (99 = matched by name only)
    reading: str
    name: str  # name_relation result

    def describe(self) -> str:
        if self.distance == 0:
            serial = "시리얼 일치"
        elif self.distance < 1:
            serial = f"시리얼 유사({self.reading}→{self.record.serial})"
        elif self.distance < 99:
            serial = f"시리얼 {int(self.distance)}자 차이({self.reading}→{self.record.serial})"
        else:
            serial = "시리얼 미확인"
        name = {"exact": "이름 일치", "partial": "이름 유사", "mismatch": "이름 불일치",
                "unknown": "이름 확인 불가"}[self.name]
        return f"{serial}, {name}"


def find_candidates(rec: Recognition, records: list[ToolRecord]) -> list[Candidate]:
    readings = []
    for value in [rec.serial_number, *rec.serial_alternatives]:
        if normalize_serial(value) and value not in readings:
            readings.append(value)
    out: list[Candidate] = []
    for record in records:
        relation = name_relation(rec.engineer_name, record)
        best = None
        for i, reading in enumerate(readings):
            d = serial_distance(reading, record.serial)
            limit = 2 if len(normalize_serial(record.serial)) >= 7 else 1
            if d > limit:
                continue
            base = {0: 100, 0.5: 95}.get(d, 70 if d == 1 else 45)
            if i > 0:
                base -= 5  # an alternative reading, not the primary one
            if best is None or base > best[0]:
                best = (base, d, reading)
        if best is None:
            if relation != "exact" or readings:
                continue
            best = (30, 99, "")  # no serial could be read: offer the engineer's tools
        score = best[0] + {"exact": 25, "partial": 10, "mismatch": -30, "unknown": 0}[relation]
        if record.disposed:
            score -= 40
        out.append(Candidate(record, score, best[1], best[2], relation))
    out.sort(key=lambda c: c.score, reverse=True)
    return out


@dataclass
class MatchDecision:
    candidate: Candidate | None
    status: str  # 'ok' | 'review' | 'error'
    messages: list[str] = field(default_factory=list)


def decide(candidates: list[Candidate]) -> MatchDecision:
    if not candidates:
        return MatchDecision(None, "error", ["엑셀에서 일치하는 툴(시리얼/이름)을 찾지 못했습니다. 대상 행을 직접 선택하세요."])
    top = candidates[0]
    runner_up = candidates[1] if len(candidates) > 1 else None
    msgs = [f"매칭: {top.record.label()} ({top.describe()})"]
    clear_lead = runner_up is None or runner_up.score <= top.score - 20

    if top.record.disposed:
        return MatchDecision(top, "review", msgs + ["폐기 처리된 툴입니다. 확인 후 반영하세요."])
    if top.distance <= 0.5 and top.name in ("exact", "unknown") and clear_lead:
        if top.distance == 0.5:
            msgs.append("시리얼 표기 차이(0/O, 앞자리 0 등)는 자동 보정했습니다.")
        return MatchDecision(top, "ok", msgs)
    if top.distance == 1 and top.name == "exact" and clear_lead:
        msgs.append("시리얼 한 글자가 다르지만 이름이 일치해 자동 보정했습니다.")
        return MatchDecision(top, "ok", msgs)

    if top.name == "mismatch":
        msgs.append("사진의 이름과 엑셀 담당자가 다릅니다.")
    if top.distance >= 99:
        msgs.append("시리얼을 읽지 못해 이름으로만 찾았습니다.")
    elif top.distance >= 1:
        msgs.append("시리얼이 정확히 일치하지 않습니다.")
    if not clear_lead:
        msgs.append("비슷한 후보가 여러 개입니다: " + " / ".join(c.record.label() for c in candidates[1:3]))
    return MatchDecision(top, "review", msgs)


# ---------------------------------------------------------------- dates
def parse_date(value) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = str(value or "").strip()
    if not text:
        return None
    m = re.search(r"(\d{4})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})", text)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", text)  # M/D/YYYY (SMAX)
        if m:
            mo, d, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        else:
            m = re.search(r"\b(\d{2})[.\-](\d{1,2})[.\-](\d{1,2})\b", text)  # YY.MM.DD
            if not m:
                return None
            y, mo, d = 2000 + int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return dt.date(y, mo, d)
    except ValueError:
        return None


def add_months(value: dt.date, months: int) -> dt.date:
    """Same as Excel EDATE: month arithmetic clamped to the month's last day."""
    year, month = divmod(value.year * 12 + value.month - 1 + months, 12)
    month += 1
    return dt.date(year, month, min(value.day, calendar.monthrange(year, month)[1]))
