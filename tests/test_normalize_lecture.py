# -*- coding: utf-8 -*-
"""normalize_lecture 강좌 단위 계산·대상 월 선별·연도 롤오버 회귀 테스트.

시트 1장 = 한 번에 결제하는 강좌 1개. 회차·기간·수강료는 시트 전체로 계산하고(달력 월로
자르지 않음), 월별 계획서는 '그 달 강좌인 정규반'만 고른다. 날짜 fixture는 실제 강좌
기간(개강~종강)에서 가져왔다.
반환 형태: normalize_lecture(...) -> (lecture | None, reports)
"""
import pytest

try:
    from src.normalize import normalize_lecture, season_months
except ModuleNotFoundError:
    from ..src.normalize import normalize_lecture, season_months


def make_raw(dates, gubun="정규반", season="", subject="국어"):
    """진도 날짜 리스트로 합성 강좌 raw 생성(각 행 '수업 주제'=N강).
    요일 표기 없이 'M/D'만 써서 요일 불일치 잡음을 피한다."""
    fields = {
        "강사명": "홍길동", "구분": gubun, "강의형태": "현장 강의",
        "과목": subject, "수업 요일 / 시간": "화 19:00",
    }
    if season:
        fields["시즌"] = season
    return {
        "source_file": "t.xlsx", "source_sheet": "강좌1",
        "fields": fields,
        "progress": [{"날짜": d, "수업 주제": f"{i + 1}강"} for i, d in enumerate(dates)],
    }


# ── 전체 모드(슬라이싱 없음) ──

def test_full_mode_counts_all_months():
    """target_month 없으면 전 기간 전 회차를 센다."""
    lec, _ = normalize_lecture(make_raw(["7/7", "7/14", "7/21", "7/28", "8/4", "8/11"]), 1, 2026)
    assert lec is not None
    assert lec["computed_total_sessions"] == 6
    assert lec["first_class_date"] == "2026-07-07"
    assert lec["last_class_date"] == "2026-08-11"


# ── 강좌 단위 계산(달력 월로 자르지 않음) ──

def codes(reports):
    return [r["issue_code"] for r in reports]


def test_block_crossing_month_end_is_billed_whole():
    """10월 강좌 10/3~11/14(10/10 휴강): 6회 전부가 회차·수강료에 들어간다.
    이전에는 10월분 4회만 계산돼 수강료가 적게 나왔다."""
    raw = make_raw(["10/3", "10/10", "10/17", "10/24", "10/31", "11/7", "11/14"], season="10월")
    raw["progress"][1]["비고"] = "휴강"
    lec, _ = normalize_lecture(raw, 1, 2026, target_month=10)
    assert lec["computed_total_sessions"] == 6
    assert lec["first_class_date"] == "2026-10-03"
    assert lec["last_class_date"] == "2026-11-14"
    assert lec["computed_fee"] == 6 * 70000
    assert lec["fee_display"] == "420,000원 (회차당 70,000원)"
    assert len(lec["progress"]) == 7                       # 진도 표시는 7행 그대로(잘라내지 않음)


def test_full_mode_and_target_mode_give_same_fee():
    """대상 월을 주든 안 주든 같은 강좌의 수강료는 같다."""
    dates = ["8/24", "8/31", "9/7", "9/14", "9/21", "9/28"]
    full, _ = normalize_lecture(make_raw(dates, season="8월"), 1, 2026)
    monthly, _ = normalize_lecture(make_raw(dates, season="8월"), 1, 2026, target_month=8)
    assert full["computed_fee"] == monthly["computed_fee"] == 6 * 70000


# ── 대상 월 선별: 시즌 칸 우선, 없으면 개강일 ──

def test_season_decides_target_month():
    raw = make_raw(["8/24", "8/31", "9/7", "9/14", "9/21", "9/28"], season="8월")
    assert normalize_lecture(raw, 1, 2026, target_month=8)[0] is not None
    lec, reports = normalize_lecture(raw, 1, 2026, target_month=9)
    assert lec is None
    assert codes(reports) == ["TARGET_MONTH_EXCLUDED"]     # 조용히 빠지지 않고, 그 리포트만 남는다
    assert reports[0]["severity"] == "정보"


def test_season_with_two_months_is_included_in_both():
    raw = make_raw(["7/7", "7/14", "7/21", "7/28", "8/4", "8/11"], season="7,8월")
    for month in (7, 8):
        lec, _ = normalize_lecture(raw, 1, 2026, target_month=month)
        assert lec is not None and lec["computed_total_sessions"] == 6


def test_start_month_decides_when_season_has_no_month():
    raw = make_raw(["7/7", "7/14", "7/21", "7/28"])
    assert normalize_lecture(raw, 1, 2026, target_month=7)[0] is not None
    lec, reports = normalize_lecture(raw, 1, 2026, target_month=9)
    assert lec is None
    assert reports[0]["issue_code"] == "TARGET_MONTH_EXCLUDED"
    assert reports[0]["severity"] == "정보"


def test_late_month_start_without_season_asks_for_review():
    """12/27 개강 1월 강좌인데 시즌 칸이 비면 개강월(12월)로 판별된다 — 어느 쪽 실행에서도 확인 요청."""
    raw = make_raw(["12/27", "1/3", "1/10", "1/17", "1/24", "1/31"])
    lec, reports = normalize_lecture(raw, 1, 2027, target_month=1)
    assert lec is None
    assert reports[0]["issue_code"] == "TARGET_MONTH_EXCLUDED"
    assert reports[0]["severity"] == "확인필요"            # 수업이 1월에 걸쳐 있어 잘못 빠졌을 수 있음
    lec, reports = normalize_lecture(raw, 1, 2026, target_month=12)
    assert lec is not None
    assert "TARGET_MONTH_AMBIGUOUS" in codes(reports)      # 넣긴 했지만 수업 대부분이 1월


def test_wrong_year_date_cells_are_excluded_with_report():
    """작년 날짜 셀(2025-07-07)이 든 정규반은 7월 계획서에서 빠지되 리포트가 남는다."""
    lec, reports = normalize_lecture(make_raw(["2025-07-07", "2025-07-14"]), 1, 2026, target_month=7)
    assert lec is None
    assert reports[0]["issue_code"] == "TARGET_MONTH_EXCLUDED"
    assert reports[0]["computed_value"] == "2025-07-07"


def test_unreadable_dates_are_excluded_with_review_report():
    lec, reports = normalize_lecture(make_raw(["첫째주", "둘째주"], season="7월"), 1, 2026, target_month=7)
    assert lec is None
    assert reports[0]["issue_code"] == "TARGET_MONTH_EXCLUDED"
    assert reports[0]["severity"] == "확인필요"


@pytest.mark.parametrize(
    "season, expected",
    [
        ("10월", [10]), ("7,8월", [7, 8]), ("7월, 8월", [7, 8]), ("11~1월", [11, 12, 1]),
        ("2026년 7월", [7]), ("5월 2주차", [5]), ("썸머", []), ("", []), (None, []), ("13월", []),
    ],
)
def test_season_months(season, expected):
    assert season_months(season) == expected


# ── 12월 → 1월 연도 롤오버 ──

def test_year_rollover_full_mode():
    """12월→1월 진도는 1월을 다음 해로 보정해 기간이 역전되지 않는다."""
    raw = make_raw(["12/7", "12/14", "12/28", "1/4", "1/11"])
    lec, _ = normalize_lecture(raw, 1, 2026)
    assert lec["computed_total_sessions"] == 5
    assert lec["first_class_date"] == "2026-12-07"
    assert lec["last_class_date"] == "2027-01-11"          # 1/11이 2026이 아니라 2027
    jan_rows = [r for r in lec["progress"] if r["parsed_date"].startswith("2027-01")]
    assert len(jan_rows) == 2


def test_january_block_starting_in_late_december():
    """2027년 1월 계획서: 12/27 개강 강좌의 12/27은 2026년으로 읽는다(실데이터: 1월 강좌 16개가 12/27~31 개강).
    이전에는 base_year를 2026으로 주든 2027로 주든 이 강좌가 통째로 사라졌다."""
    raw = make_raw(["12/27", "1/3", "1/10", "1/17", "1/24", "1/31"], season="1월")
    lec, reports = normalize_lecture(raw, 1, 2027, target_month=1)
    assert lec is not None
    assert lec["first_class_date"] == "2026-12-27"
    assert lec["last_class_date"] == "2027-01-31"
    assert lec["computed_total_sessions"] == 6
    assert lec["computed_fee"] == 6 * 70000
    assert "OUT_OF_ORDER_DATE" not in codes(reports)


def test_december_block_keeps_base_year():
    raw = make_raw(["11/24", "12/1", "12/8", "12/15"], season="12월")
    lec, _ = normalize_lecture(raw, 1, 2026, target_month=12)
    assert lec["first_class_date"] == "2026-11-24"


# ── 특강은 월 모드라도 항상 포함 ──

def test_special_lecture_always_included_in_monthly_mode():
    """특강(썸머/윈터)은 어느 달을 뽑든 전체 기간 그대로 포함."""
    raw = make_raw(["7/7", "7/14", "8/4", "8/11"], gubun="특강", season="썸머")
    for month in (7, 11):
        lec, _ = normalize_lecture(raw, 1, 2026, target_month=month)
        assert lec is not None
        assert lec["computed_total_sessions"] == 4
        assert lec["billing"] == "total"


# ── 휴강 행은 회차에서 제외 ──

def test_holiday_row_not_counted():
    """진도표의 휴강 행은 실제 회차(total_sessions)에 포함되지 않는다."""
    raw = {
        "source_file": "t.xlsx", "source_sheet": "강좌1",
        "fields": {"강사명": "홍길동", "구분": "정규반", "강의형태": "현장 강의",
                   "과목": "국어", "수업 요일 / 시간": "화 19:00"},
        "progress": [
            {"날짜": "7/7", "수업 주제": "1강"},
            {"날짜": "7/14", "수업 주제": "2강"},
            {"날짜": "7/21", "수업 주제": "휴강"},
            {"날짜": "7/28", "수업 주제": "3강"},
        ],
    }
    lec, _ = normalize_lecture(raw, 1, 2026)
    assert lec["computed_total_sessions"] == 3             # 휴강 1건 제외
    assert len(lec["progress"]) == 4                       # 진도 표시는 4행 그대로
