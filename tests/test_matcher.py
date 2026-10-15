import datetime as dt

from toolcal.matcher import (
    add_months,
    decide,
    find_candidates,
    name_relation,
    normalize_serial,
    parse_date,
    serial_distance,
)
from toolcal.recognize import Recognition
from toolcal.workbook import ToolWorkbook


def _rec(name="", serial="", alts=()):
    return Recognition(image_type="tool_photo", engineer_name=name, serial_number=serial,
                       serial_alternatives=list(alts))


def test_serial_normalisation():
    assert normalize_serial(" 4435-0259 sv ") == "44350259SV"
    assert serial_distance("31550075", "31550075") == 0
    assert serial_distance("716805829", "0716805829") == 0.5  # leading zero lost
    assert serial_distance("3155OO75", "31550075") == 0.5  # letter O for zero
    assert serial_distance("31550076", "31550075") == 1


def test_exact_match_is_ok(workbook_path):
    wb = ToolWorkbook.open(workbook_path)
    d = decide(find_candidates(_rec("박민수", "31550075"), wb.records))
    assert d.status == "ok"
    assert (d.candidate.record.sheet, d.candidate.record.row) == ("광주", 3)


def test_name_picks_between_close_serials(workbook_path):
    wb = ToolWorkbook.open(workbook_path)
    # one digit misread; 31550069 (이영희) is 2 edits away, 31550075 (박민수) 1 edit
    d = decide(find_candidates(_rec("박민수", "31550078"), wb.records))
    assert d.status == "ok" and d.candidate.record.serial == "31550075"


def test_fuzzy_without_name_needs_review(workbook_path):
    wb = ToolWorkbook.open(workbook_path)
    d = decide(find_candidates(_rec("", "31550078"), wb.records))
    assert d.status == "review" and d.candidate.record.serial == "31550075"


def test_name_mismatch_needs_review(workbook_path):
    wb = ToolWorkbook.open(workbook_path)
    d = decide(find_candidates(_rec("이영희", "31550075"), wb.records))
    assert d.status == "review"
    assert any("이름" in m for m in d.messages)


def test_alternative_reading_is_used(workbook_path):
    wb = ToolWorkbook.open(workbook_path)
    d = decide(find_candidates(_rec("김철수", "46900064", alts=["46900004"]), wb.records))
    assert d.candidate.record.serial == "46900004" and d.status == "ok"


def test_unknown_serial_is_error(workbook_path):
    wb = ToolWorkbook.open(workbook_path)
    assert decide(find_candidates(_rec("", "99999999"), wb.records)).status == "error"


def test_name_only_offers_engineers_tools(workbook_path):
    wb = ToolWorkbook.open(workbook_path)
    cands = find_candidates(_rec("김철수", ""), wb.records)
    assert {c.record.serial for c in cands} == {"218800584", "46900004"}
    assert decide(cands).status == "review"


def test_disposed_tool_needs_review(workbook_path):
    wb = ToolWorkbook.open(workbook_path)
    d = decide(find_candidates(_rec("박민수", "0116503968"), wb.records))
    assert d.status == "review" and d.candidate.record.disposed


def test_name_relation(workbook_path):
    wb = ToolWorkbook.open(workbook_path)
    rec = next(r for r in wb.records if r.serial == "31550075")
    assert name_relation("박민수", rec) == "exact"
    assert name_relation("박민주", rec) == "partial"
    assert name_relation("홍길동", rec) == "mismatch"
    assert name_relation("MINSU PARK", rec) == "unknown"


def test_dates():
    assert parse_date("2025.12.29") == dt.date(2025, 12, 29)
    assert parse_date("2025-12-29") == dt.date(2025, 12, 29)
    assert parse_date("12/29/2026") == dt.date(2026, 12, 29)
    assert parse_date("25.12.29") == dt.date(2025, 12, 29)
    assert parse_date("") is None and parse_date("2025.13.40") is None
    assert add_months(dt.date(2025, 12, 29), 12) == dt.date(2026, 12, 29)
    assert add_months(dt.date(2025, 8, 31), 6) == dt.date(2026, 2, 28)
    assert add_months(dt.date(2026, 6, 24), -6) == dt.date(2025, 12, 24)
