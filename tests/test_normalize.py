# -*- coding: utf-8 -*-
"""classify_billing 정규/특강 청구 판정 회귀 테스트.

월(monthly) vs 전체(total) 오판은 수강료 과다/과소청구로 직결된다.
이번 프로젝트에서 두 번 고친 규칙(정규 우선 처리, 썸머/윈터 대칭 인식)을
경계 케이스로 고정해 회귀를 막는다.
"""
import pytest

try:
    from src.normalize import classify_billing
except ModuleNotFoundError:
    from ..src.normalize import classify_billing


@pytest.mark.parametrize(
    "fields, expected",
    [
        # 단독 — 가장 기본
        ({"구분": "정규반"}, (True, "monthly")),
        ({"구분": "특강"}, (False, "total")),
        # 정규+특강 혼합 → 정규 우선(과다청구 방지)
        ({"구분": "정규반, 특강"}, (True, "monthly")),
        ({"구분": "정규특강"}, (True, "monthly")),
        # 시즌만으로 특강 인식(썸머/윈터 대칭)
        ({"시즌": "썸머"}, (False, "total")),
        ({"시즌": "윈터"}, (False, "total")),
        ({"시즌": "썸머특강"}, (False, "total")),
        ({"시즌": "윈터특강"}, (False, "total")),
        ({"시즌": "특강"}, (False, "total")),
        # 정규반이면 시즌이 윈터여도 monthly(정규 우선)
        ({"구분": "정규반", "시즌": "윈터"}, (True, "monthly")),
        ({"구분": "정규반", "시즌": "썸머"}, (True, "monthly")),
        # 구분·시즌 동시에 특강 신호
        ({"구분": "특강", "시즌": "썸머"}, (False, "total")),
        # 특강 신호 없음 → 기본 monthly(특강 아님)
        ({}, (False, "monthly")),
        ({"구분": "단과"}, (False, "monthly")),
        ({"구분": "", "시즌": ""}, (False, "monthly")),
    ],
    ids=[
        "정규반_단독",
        "특강_단독",
        "정규+특강_혼합_정규우선",
        "정규특강_연결_정규우선",
        "시즌_썸머",
        "시즌_윈터",
        "시즌_썸머특강",
        "시즌_윈터특강",
        "시즌_특강만",
        "정규+윈터_정규우선",
        "정규+썸머_정규우선",
        "특강+썸머_동시",
        "빈입력_기본monthly",
        "단과_특강아님",
        "빈문자열_기본monthly",
    ],
)
def test_classify_billing(fields, expected):
    assert classify_billing(fields) == expected


def test_regular_priority_over_special():
    """정규+특강이 동시에 붙으면 정규 우선 → monthly(과다청구 방지)."""
    assert classify_billing({"구분": "정규반, 특강"}) == (True, "monthly")
    assert classify_billing({"구분": "정규반", "시즌": "윈터"}) == (True, "monthly")


def test_winter_season_symmetric_with_summer():
    """시즌에 '특강' 단어 없이 썸머/윈터만 적어도 특강(total)으로 본다 — 대칭."""
    assert classify_billing({"시즌": "썸머"}) == (False, "total")
    assert classify_billing({"시즌": "윈터"}) == (False, "total")
