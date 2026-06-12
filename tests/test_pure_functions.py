# -*- coding: utf-8 -*-
"""순수함수 회귀 테스트 (B3): should_skip_sheet / progress_split / normalize_subject / _est_lines.

폰트·시트·과목 규칙을 바꿀 때 의도치 않은 회귀를 즉시 잡기 위한 경계 케이스 고정.
"""
import pytest

try:
    from src.collect_excel import should_skip_sheet
    from src.generate_pptx import normalize_subject, progress_split
    from src.reflow import _est_lines
except ModuleNotFoundError:
    from ..src.collect_excel import should_skip_sheet
    from ..src.generate_pptx import normalize_subject, progress_split
    from ..src.reflow import _est_lines


# ── should_skip_sheet: 이름·상태 기준, 위치 무관 ──

class _FakeWS:
    def __init__(self, title, sheet_state="visible"):
        self.title = title
        self.sheet_state = sheet_state


@pytest.mark.parametrize(
    "title, state, expected",
    [
        ("강좌1", "visible", False),
        ("강좌2", "visible", False),
        ("물리 김선생", "visible", False),   # 이름 바꿔도 강좌로 인식(위치·이름 무관)
        ("수학심화반", "visible", False),
        ("작성 안내", "visible", True),       # '안내'
        ("사용 안내", "visible", True),       # '안내' 부분일치(이름 바뀌어도 제외)
        ("예시_통합", "visible", True),       # '예시'
        ("샘플강좌", "visible", True),        # '샘플'
        ("선택목록", "hidden", True),         # 숨김 + 키워드
        ("강좌1", "hidden", True),            # 숨김이면 강좌명이어도 제외
        ("선택목록", "visible", True),        # 이름 키워드로도 제외
    ],
)
def test_should_skip_sheet(title, state, expected):
    assert should_skip_sheet(_FakeWS(title, state)) is expected


def test_should_skip_sheet_defaults_visible_when_no_state():
    """sheet_state 속성이 없으면 visible로 간주(이름이 강좌면 처리)."""
    class _NoState:
        title = "강좌1"
    assert should_skip_sheet(_NoState()) is False


# ── progress_split: 좌우 균형 분배(좌측 먼저, 각 최대 5, 합 최대 10) ──

@pytest.mark.parametrize(
    "n, expected",
    [
        (0, (0, 0)),
        (1, (1, 0)),
        (2, (1, 1)),
        (3, (2, 1)),
        (4, (2, 2)),
        (5, (3, 2)),
        (6, (3, 3)),
        (9, (5, 4)),
        (10, (5, 5)),
        (11, (5, 5)),    # 10개 초과는 10개로 캡
        (100, (5, 5)),
    ],
)
def test_progress_split(n, expected):
    assert progress_split(n) == expected


def test_progress_split_invariants():
    """모든 n에서 좌·우 ≤ 5, 좌측 먼저 채움, 합 = min(n, 10)."""
    for n in range(0, 30):
        left, right = progress_split(n)
        assert 0 <= right <= left <= 5
        assert left + right == min(n, 10)


# ── normalize_subject: 수학 통일 + 수준표기 제거 ──

@pytest.mark.parametrize(
    "subject, expected",
    [
        ("미적분", "수학"),
        ("확률과 통계", "수학"),
        ("기하", "수학"),
        ("수학", "수학"),
        ("수리논술", "수리논술"),   # 핵심: '수리'는 수학 키가 아니라 그대로 둔다(보라 테두리)
        ("논술", "논술"),
        ("물리학Ⅰ", "물리학"),       # 로마자 수준표기 제거
        ("화학1", "화학"),           # 아라비아 제거
        ("통합과학2", "통합과학"),
        ("생명과학Ⅱ", "생명과학"),
        ("물리학１", "물리학"),       # 전각 숫자 제거
        ("국어", "국어"),
        ("  영어  ", "영어"),
        ("", ""),
        (None, ""),
    ],
)
def test_normalize_subject(subject, expected):
    assert normalize_subject(subject) == expected


# ── _est_lines: 줄 수 추정 ──

@pytest.mark.parametrize(
    "text, width_pt, font_pt, wrap, expected",
    [
        ("", 100, 10, True, 0),               # 빈 텍스트 → 0
        (None, 100, 10, True, 0),
        ("가나다", 1000, 10, True, 1),         # 짧고 넓음 → 1줄
        ("a\nb\nc", 5, 10, False, 3),         # wrap off → 명시 줄바꿈 수(가로 무시)
        ("oneline", 5, 10, False, 1),
        ("a\nb", 1000, 10, True, 2),          # 짧은 두 줄
        ("hi\n\n\n", 1000, 10, True, 1),      # 끝 공백·줄바꿈 제거 후 1줄
    ],
)
def test_est_lines_structural(text, width_pt, font_pt, wrap, expected):
    assert _est_lines(text, width_pt, font_pt, wrap) == expected


@pytest.mark.parametrize(
    "text, width_pt, font_pt, expected",
    [
        # 현재 폭 비율(한글 GLYPH_W_KO=0.88 / ASCII 0.61) 기준 — 비율 바꾸면 이 값 갱신(회귀 잠금).
        ("가나다라마", 20, 10, 3),   # 5*8.8=44.0 → ceil(44/20)=3
        ("abcde", 20, 10, 2),        # 5*6.1=30.5 → ceil(30.5/20)=2
    ],
)
def test_est_lines_wrapping_locks_ratio(text, width_pt, font_pt, expected):
    assert _est_lines(text, width_pt, font_pt, wrap=True) == expected
