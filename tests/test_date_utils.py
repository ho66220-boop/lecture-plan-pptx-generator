from datetime import date

from src.date_utils import extract_start_time_display, parse_date, weekday_kr


def test_parse_slash_date_with_weekday():
    parsed, claimed = parse_date("5/4(월)", base_year=2026)
    assert parsed == date(2026, 5, 4)
    assert claimed == "월"
    assert weekday_kr(parsed) == "월"


def test_parse_korean_date():
    parsed, claimed = parse_date("5월 12일", base_year=2026)
    assert parsed == date(2026, 5, 12)
    assert claimed is None


def test_parse_iso_date():
    parsed, claimed = parse_date("2026-05-12", base_year=2026)
    assert parsed == date(2026, 5, 12)
    assert claimed is None


def test_extract_start_time_display():
    assert extract_start_time_display("매주 월요일 18:30~22:00") == "오후 6시 30분"
