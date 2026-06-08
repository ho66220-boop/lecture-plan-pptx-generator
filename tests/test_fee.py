# -*- coding: utf-8 -*-
"""calculate_fee 수강료 계산 테스트."""
try:
    from src.fee import calculate_fee
except ModuleNotFoundError:
    from ..src.fee import calculate_fee


def test_basic_table_onsite():
    """현장강의는 회차당 80,000원 × 총회차."""
    fee = calculate_fee("L001", "현장강의", 8)
    assert fee["fee_per_session"] == 80000
    assert fee["computed_fee"] == 640000
    assert fee["used_override"] is False


def test_format_key_ignores_spaces():
    """'현장 강의'처럼 띄어쓰기가 달라도 같은 단가로 매칭."""
    fee = calculate_fee("L002", "현장 강의", 4)
    assert fee["fee_per_session"] == 80000
    assert fee["computed_fee"] == 320000


def test_live_lecture_rate():
    """LIVE 강의는 회차당 50,000원."""
    fee = calculate_fee("L003", "LIVE 강의", 10)
    assert fee["fee_per_session"] == 50000
    assert fee["computed_fee"] == 500000


def test_teacher_override_beats_format():
    """강사명 예외가 있으면 강의형태 표보다 우선(형태와 무관하게 적용)."""
    onsite = calculate_fee("L004", "현장강의", 8, teacher_name="홍길동")
    live = calculate_fee("L005", "LIVE 강의", 4, teacher_name="홍길동")
    assert onsite["fee_per_session"] == 70000
    assert onsite["computed_fee"] == 560000
    assert onsite["used_override"] is True
    assert live["fee_per_session"] == 70000  # 형태가 달라도 예외 단가 동일


def test_unknown_format_returns_empty():
    """등록되지 않은 강의형태는 수강료를 비워 두고 used_override=False."""
    fee = calculate_fee("L006", "온라인특강", 6)
    assert fee["fee_display"] == ""
    assert fee["computed_fee"] is None


def test_zero_sessions_returns_empty():
    """총회차가 0이면 수강료를 계산하지 않는다."""
    fee = calculate_fee("L007", "현장강의", 0)
    assert fee["fee_display"] == ""
    assert fee["computed_fee"] is None


def test_monthly_billing_uses_month_sessions():
    """정규반(월 청구)은 한 달 회차수 기준. 총 8회·월 4회면 320,000원."""
    fee = calculate_fee("L008", "현장강의", 8, monthly_sessions=4, billing="monthly")
    assert fee["fee_per_session"] == 80000
    assert fee["computed_fee"] == 320000
    assert fee["billing"] == "monthly"
    assert fee["fee_display"].startswith("월 ")


def test_special_billing_uses_total():
    """특강/썸머(total 청구)는 전체 회차 합계."""
    fee = calculate_fee("L009", "현장강의", 4, monthly_sessions=4, billing="total")
    assert fee["computed_fee"] == 320000
    assert fee["billing"] == "total"
    assert not fee["fee_display"].startswith("월 ")


def test_monthly_with_teacher_override():
    """월 청구에도 강사 예외 단가가 적용된다(70,000 × 월 4회)."""
    fee = calculate_fee("L010", "현장강의", 8, teacher_name="홍길동", monthly_sessions=4, billing="monthly")
    assert fee["fee_per_session"] == 70000
    assert fee["computed_fee"] == 280000
