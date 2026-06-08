import unicodedata

try:
    from config.defaults import FEE_OVERRIDES, FEE_PER_SESSION_OVERRIDES, FEE_TABLE
except ModuleNotFoundError:
    from ..config.defaults import FEE_OVERRIDES, FEE_PER_SESSION_OVERRIDES, FEE_TABLE


def _normalize_format_key(value):
    """공백/대소문자 차이를 무시하기 위한 강의형태 키 정규화."""
    return "".join((value or "").split()).lower()


def _normalize_name_key(value):
    """강사명 비교용 정규화(공백 제거 + NFC)."""
    return unicodedata.normalize("NFC", "".join((value or "").split()))


# 강사가 "현장강의" / "현장 강의"처럼 띄어쓰기를 다르게 적어도 매칭되도록 정규화 키로 조회.
_FEE_TABLE_NORMALIZED = {_normalize_format_key(k): v for k, v in FEE_TABLE.items()}
_FEE_BY_TEACHER = {_normalize_name_key(k): v for k, v in FEE_PER_SESSION_OVERRIDES.items()}


def calculate_fee(
    lecture_id,
    lecture_format,
    total_sessions,
    teacher_name="",
    monthly_sessions=None,
    billing="total",
):
    """수강료 계산.

    billing="monthly": 정규반 등 월 단위 청구. 한 달 회차수(monthly_sessions) 기준으로 산출.
    billing="total":   특강/썸머 등 전체 회차 합계로 산출(강좌 전체 기간 일괄 수강).
    """
    override = FEE_OVERRIDES.get(lecture_id)
    if override is not None:
        total = int(override)
        return {
            "fee_per_session": None,
            "computed_fee": total,
            "fee_display": f"{total:,}원",
            "used_override": True,
            "billing": "override",
        }

    # 강사명 기준 회차당 예외가 있으면 강의형태 표보다 우선.
    per_session = _FEE_BY_TEACHER.get(_normalize_name_key(teacher_name))
    used_override = per_session is not None
    if per_session is None:
        per_session = _FEE_TABLE_NORMALIZED.get(_normalize_format_key(lecture_format))
    if per_session is None or not total_sessions:
        return {
            "fee_per_session": per_session,
            "computed_fee": None,
            "fee_display": "",
            "used_override": False,
            "billing": billing,
        }

    per_session = int(per_session)
    if billing == "monthly" and monthly_sessions:
        billed_sessions = int(monthly_sessions)
        total = per_session * billed_sessions
        fee_display = f"월 {total:,}원 (회차당 {per_session:,}원)"
    else:
        billing = "total"
        billed_sessions = int(total_sessions)
        total = per_session * billed_sessions
        fee_display = f"{total:,}원 (회차당 {per_session:,}원)"

    return {
        "fee_per_session": per_session,
        "computed_fee": total,
        "fee_display": fee_display,
        "used_override": used_override,
        "billing": billing,
    }
