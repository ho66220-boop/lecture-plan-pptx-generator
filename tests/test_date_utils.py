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


def test_range_like_text_not_parsed_as_date():
    # "3-4교시", "1-2회" 같은 범위 표기를 날짜로 오인하면 안 된다.
    assert parse_date("3-4교시")[0] is None
    assert parse_date("1-2회")[0] is None
    assert parse_date("5/6회차")[0] is None
    # 정상 날짜는 그대로 파싱.
    assert parse_date("7/20(월)")[0] is not None


def test_date_before_hoe_word_not_corrupted():
    # '회'로 시작하는 일반 단어(회의·회식)가 뒤따라도 날짜는 온전히 파싱돼야 한다.
    # 단위어 회피가 백트래킹을 유발하면 "3/15 회의"가 3/1로 왜곡되는 회귀가 있었다(잘림 금지).
    assert parse_date("3/15 회의실 변경", base_year=2026)[0] == date(2026, 3, 15)
    assert parse_date("9/1 회식", base_year=2026)[0] == date(2026, 9, 1)
    # 날짜와 범위 표기가 섞이면 날짜 쪽을 잡는다.
    assert parse_date("7/21 3-4교시 휴강", base_year=2026)[0] == date(2026, 7, 21)
