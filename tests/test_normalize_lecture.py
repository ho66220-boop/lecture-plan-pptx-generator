# -*- coding: utf-8 -*-
"""normalize_lecture 월 슬라이싱·연도 롤오버·다음 달 미리보기·0회 제외 회귀 테스트.

월별 계획서는 청구·회차에 직결되고, 12월→1월 연도 경계는 이번에 고친 미묘한
버그라 회귀 위험이 크다. 합성 진도 fixture로 경계 동작을 고정한다.
반환 형태: normalize_lecture(...) -> (lecture | None, reports)
"""
try:
    from src.normalize import normalize_lecture
except ModuleNotFoundError:
    from ..src.normalize import normalize_lecture


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


# ── 월 슬라이싱 ──

def test_target_month_slices_to_that_month():
    """target_month=7이면 7월 회차만 청구에 반영(8월은 다음 달 미리보기)."""
    raw = make_raw(["7/7", "7/14", "7/21", "7/28", "8/4", "8/11"])
    lec, _ = normalize_lecture(raw, 1, 2026, target_month=7)
    assert lec["computed_total_sessions"] == 4           # 7월만 회차 집계
    assert lec["first_class_date"] == "2026-07-07"
    assert lec["last_class_date"] == "2026-07-28"
    this_month = [r for r in lec["progress"] if not r["is_next_month"]]
    next_month = [r for r in lec["progress"] if r["is_next_month"]]
    assert len(this_month) == 4
    assert len(next_month) == 2                            # 8월 행은 미리보기로 함께 표시
    assert all(not r["is_real_class"] for r in next_month) # 미리보기는 회차에 안 들어감


def test_target_month_excludes_other_months():
    """그 달·다음 달이 아닌 행은 진도에서 제외된다(8월: 7월 행 제거, 9월 없음)."""
    raw = make_raw(["7/7", "7/14", "7/21", "7/28", "8/4", "8/11"])
    lec, _ = normalize_lecture(raw, 1, 2026, target_month=8)
    assert lec["computed_total_sessions"] == 2
    assert all(r["parsed_date"].startswith("2026-08") for r in lec["progress"])


def test_zero_class_month_returns_none():
    """그 달에 수업이 0회인 정규반은 (None, reports)을 돌려 슬라이드를 안 만든다."""
    raw = make_raw(["7/7", "7/14", "8/4"])
    lec, reports = normalize_lecture(raw, 1, 2026, target_month=9)
    assert lec is None
    assert isinstance(reports, list)


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


def test_year_rollover_target_december():
    """target_month=12: 12월은 그 달, 1월은 다음 해 1월 미리보기로 잡힌다."""
    raw = make_raw(["12/7", "12/14", "12/28", "1/4", "1/11"])
    lec, _ = normalize_lecture(raw, 1, 2026, target_month=12)
    assert lec["computed_total_sessions"] == 3             # 12월만 회차
    this_month = [r for r in lec["progress"] if not r["is_next_month"]]
    next_month = [r for r in lec["progress"] if r["is_next_month"]]
    assert len(this_month) == 3
    assert len(next_month) == 2
    assert all(r["parsed_date"].startswith("2027-01") for r in next_month)


# ── 특강은 월 모드라도 슬라이싱하지 않음 ──

def test_special_lecture_not_sliced_in_monthly_mode():
    """특강(썸머/윈터)은 패키지 단위라 target_month여도 전체 기간 그대로 포함(total 청구)."""
    raw = make_raw(["7/7", "7/14", "8/4", "8/11"], gubun="특강", season="썸머")
    lec, _ = normalize_lecture(raw, 1, 2026, target_month=7)
    assert lec is not None
    assert lec["computed_total_sessions"] == 4             # 슬라이싱 안 함(8월도 포함)
    assert lec["billing"] == "total"
    assert all(not r["is_next_month"] for r in lec["progress"])  # 미리보기 표식 없음


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
