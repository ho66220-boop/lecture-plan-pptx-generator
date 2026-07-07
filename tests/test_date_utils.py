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


def test_excel_serial_date_converted():
    # 날짜 서식이 풀려 숫자(엑셀 일련번호)로 저장된 셀도 날짜로 복구한다.
    from datetime import timedelta
    d = parse_date("46225", base_year=2026)[0]
    assert d == date(1899, 12, 30) + timedelta(days=46225)   # 1900 날짜 시스템 기준
    assert d.year == 2026
    # 7일 간격 연속 회차(46225 → 46232)도 정확히 일주일 차이.
    assert (parse_date("46232")[0] - parse_date("46225")[0]).days == 7
    # float/시간분수로 들어와도 그날로.
    assert parse_date("46225.0")[0] == d
    assert parse_date("46225.75")[0] == d


def test_non_date_numbers_not_converted():
    # 회차 수·연도 단독·구분자 없는 오타 등 날짜 아닌 숫자는 변환하지 않는다(오변환 방지).
    assert parse_date("3")[0] is None
    assert parse_date("12")[0] is None
    assert parse_date("2026")[0] is None        # 연도만 단독 → 날짜 아님(범위 밖)
    assert parse_date("20260706")[0] is None    # 구분자 없는 8자리 → 변환 안 함
    # 구분자 있는 정상 날짜는 그대로 파싱.
    assert parse_date("7/6")[0] == date(2026, 7, 6)


def test_date_before_hoe_word_not_corrupted():
    # '회'로 시작하는 일반 단어(회의·회식)가 뒤따라도 날짜는 온전히 파싱돼야 한다.
    # 단위어 회피가 백트래킹을 유발하면 "3/15 회의"가 3/1로 왜곡되는 회귀가 있었다(잘림 금지).
    assert parse_date("3/15 회의실 변경", base_year=2026)[0] == date(2026, 3, 15)
    assert parse_date("9/1 회식", base_year=2026)[0] == date(2026, 9, 1)
    # 날짜와 범위 표기가 섞이면 날짜 쪽을 잡는다.
    assert parse_date("7/21 3-4교시 휴강", base_year=2026)[0] == date(2026, 7, 21)
