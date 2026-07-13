# -*- coding: utf-8 -*-
"""Batch 2 — 돈 직결 침묵 오염 회귀 테스트 (P2-1 순서 어긋남 · P3-1 빈 날짜 행).

정책: 애매하면 값을 조용히 고치지 말고 리포트한다. 진짜 연도 경계(12→1 인접 wrap)만
연도를 올리고, 그 외 월 역행(7월 사이의 6월 보강행 등)은 base_year를 유지한 채
OUT_OF_ORDER_DATE로 사람에게 알린다. 내용이 있는데 날짜만 빈 행은
PROGRESS_DATE_MISSING으로 드러낸다(회차 산정 정책은 불변 — 날짜 없는 행 제외 유지).
"""
try:
    from src.normalize import normalize_lecture
except ModuleNotFoundError:
    from ..src.normalize import normalize_lecture


def make_raw(rows):
    """rows: [{"날짜":..., "수업 주제":..., ...}] 합성 강좌. 요일 표기는 쓰지 않아
    요일 불일치(WEEKDAY_MISMATCH) 경로와 분리해 검증한다."""
    return {
        "source_file": "t.xlsx", "source_sheet": "강좌1",
        "fields": {"강사명": "김강사", "강좌명": "테스트", "구분": "정규반",
                   "강의형태": "현장강의", "과목": "국어", "수업 요일 / 시간": "화 19:00"},
        "progress": rows,
    }


def dated(dates):
    return [{"날짜": d, "수업 주제": f"{i + 1}강"} for i, d in enumerate(dates)]


# ══ T1 — P2-1: 순서 어긋남(보강행 삽입)이 연도를 오염시키지 않는다 ══

def test_out_of_order_date_does_not_bump_year():
    """# P2-1 회귀 방지
    요일 표기 없는 진도 7/1,7/8,6/24(보강),7/15,7/22 — 현재는 6/24에서 연도가
    +1되어 이후 전 행이 2027로 밀리고 수강료가 반토막(월 160,000)·리포트 0건.
    기대: 전 행 2026 유지, 총회차 5, 월 회차 분할집계 없음(월 320,000원),
    OUT_OF_ORDER_DATE 리포트 존재."""
    lec, reports = normalize_lecture(make_raw(dated(["7/1", "7/8", "6/24", "7/15", "7/22"])), 1, 2026)

    assert lec["last_class_date"] == "2026-07-22"          # 연도 오염 없음
    assert lec["computed_total_sessions"] == 5             # 보강행도 실제 수업으로 포함
    assert lec["computed_fee"] == 320000                   # 월 회차 4 유지(분할집계 방지)
    assert lec["fee_display"] == "월 320,000원 (회차당 80,000원)"
    years = {r["parsed_date"][:4] for r in lec["progress"] if r["parsed_date"]}
    assert years == {"2026"}                               # 전 행 base_year 유지
    codes = [r["issue_code"] for r in reports]
    assert "OUT_OF_ORDER_DATE" in codes                    # 침묵이 아니라 리포트


# ══ T2 — P3-1: 내용 있는 빈 날짜 행은 리포트로 드러난다 ══

def test_empty_date_with_content_reported():
    """# P3-1 회귀 방지
    진도 3행 중 가운데 행이 날짜=""(셀 병합 소거 등) + 수업 주제="2강".
    현재는 리포트 0건으로 회차에서 조용히 누락(총회차 2 = 수강료 과소).
    기대: PROGRESS_DATE_MISSING 리포트 존재. 회차 정책은 불변(날짜 없는 행 제외
    유지 → 총회차 2) — 리포트를 보고 사람이 입력을 고쳐 재실행하는 흐름."""
    rows = [
        {"날짜": "7/7", "수업 주제": "1강"},
        {"날짜": "", "수업 주제": "2강"},
        {"날짜": "7/21", "수업 주제": "3강"},
    ]
    lec, reports = normalize_lecture(make_raw(rows), 1, 2026)

    codes = [r["issue_code"] for r in reports]
    assert "PROGRESS_DATE_MISSING" in codes                # 침묵 누락 금지
    assert lec["computed_total_sessions"] == 2             # 회차 산정 정책 불변


# ══ T3 — 경계: 완전 빈 행은 기존대로 무리포트 (동작 잠금) ══

def test_fully_empty_row_stays_silent():
    """# P3-1 경계 잠금
    내용까지 완전히 빈 행(날짜·회차·주제·상세·비고 모두 빈값)은 기존대로
    리포트 없이 무시된다 — 템플릿 잔여 빈 행에 오탐하지 않는다."""
    rows = [
        {"날짜": "7/7", "수업 주제": "1강"},
        {"날짜": "", "회차": "", "수업 주제": "", "상세 내용": "", "비고": ""},
        {"날짜": "7/21", "수업 주제": "3강"},
    ]
    lec, reports = normalize_lecture(make_raw(rows), 1, 2026)

    codes = [r["issue_code"] for r in reports]
    assert "PROGRESS_DATE_MISSING" not in codes
    assert "DATE_PARSE_FAILED" not in codes
    assert lec["computed_total_sessions"] == 2


# ══ T4 — 회귀 방지: 정상 12→1 연도 경계는 그대로 wrap + 신규 코드 미발동 ══

def test_genuine_december_wrap_still_bumps_year_without_report():
    """# P2-1 회귀 방지(반대 방향)
    정상 겨울 강좌(12/7,12/14,12/28,1/4,1/11)는 기존대로 1월을 다음 해로
    올리고, OUT_OF_ORDER_DATE는 발동하지 않아야 한다 — 판별을 좁히다
    진짜 연도 경계를 죽이면 기존 테스트 2건과 함께 여기서도 잡힌다."""
    lec, reports = normalize_lecture(make_raw(dated(["12/7", "12/14", "12/28", "1/4", "1/11"])), 1, 2026)

    assert lec["last_class_date"] == "2027-01-11"          # wrap 보존
    assert lec["computed_total_sessions"] == 5
    codes = [r["issue_code"] for r in reports]
    assert "OUT_OF_ORDER_DATE" not in codes                # 정상 경계에 오탐 없음
    assert "PROGRESS_DATE_MISSING" not in codes
