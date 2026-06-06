try:
    from config.defaults import FEE_OVERRIDES, FEE_TABLE
except ModuleNotFoundError:
    from ..config.defaults import FEE_OVERRIDES, FEE_TABLE


def _normalize_format_key(value):
    """공백/대소문자 차이를 무시하기 위한 강의형태 키 정규화."""
    return "".join((value or "").split()).lower()


# 강사가 "현장강의" / "현장 강의"처럼 띄어쓰기를 다르게 적어도 매칭되도록 정규화 키로 조회.
_FEE_TABLE_NORMALIZED = {_normalize_format_key(k): v for k, v in FEE_TABLE.items()}


def calculate_fee(lecture_id, lecture_format, total_sessions):
    override = FEE_OVERRIDES.get(lecture_id)
    if override is not None:
        per_session = None
        total = int(override)
        return {
            "fee_per_session": per_session,
            "computed_fee": total,
            "fee_display": f"{total:,}원",
            "used_override": True,
        }

    per_session = _FEE_TABLE_NORMALIZED.get(_normalize_format_key(lecture_format))
    if per_session is None or not total_sessions:
        return {
            "fee_per_session": per_session,
            "computed_fee": None,
            "fee_display": "",
            "used_override": False,
        }

    total = int(per_session) * int(total_sessions)
    return {
        "fee_per_session": int(per_session),
        "computed_fee": total,
        "fee_display": f"{total:,}원 (회차당 {int(per_session):,}원)",
        "used_override": False,
    }
