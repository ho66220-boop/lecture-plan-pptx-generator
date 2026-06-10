# -*- coding: utf-8 -*-
"""과목에 따라 슬라이드 테두리에 색 띠(프레임)를 두른다.

국어=빨강 / 수학=노랑 / 영어=주황 / 사탐=초록 / 과탐=파랑 / 논술=보라.
세부 과목명(생명과학·물리학·생활과윤리·수리논술 등)도 키워드로 6개 계열에 매핑한다.
"""
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Pt

EMU_PER_CM = 360000.0

# (계열 키워드들, 색상). 위에서부터 먼저 매칭 — '논술'을 '수학'보다 앞에 둬 수리논술을 보라로 분류.
# 출력물 과목 분류용 — 뮤트 톤으로 차분하게. 단 흰 종이 인쇄 시 가장 안 보이는 노랑(수학)만 덜 연하게.
SUBJECT_RULES = [
    (("국어",), "FF8787"),                                              # 빨강(연)
    (("영어",), "FFA94D"),                                              # 주황(연)
    (("논술", "수리논술"), "B197FC"),                                    # 보라(연)
    (("수학", "수리"), "FFD43B"),                                        # 노랑(인쇄 가독성 위해 덜 연하게)
    (("사탐", "사회", "윤리", "한국사", "지리", "경제", "정치", "역사", "세계사", "동아시아"), "69DB7C"),   # 초록(연)
    (("과탐", "과학", "물리", "화학", "생명", "지구"), "74C0FC"),          # 하늘(연)
]


def subject_color(subject):
    """과목명 → 색상(hex) 또는 None(미매칭)."""
    key = (subject or "").replace(" ", "")
    if not key:
        return None
    for keywords, rgb in SUBJECT_RULES:
        if any(word in key for word in keywords):
            return rgb
    return None


def add_subject_band(slide, subject, slide_width, slide_height, line_pt=6, inset_cm=0.18):
    """슬라이드 가장자리에 과목 색 테두리 프레임을 추가. 추가했으면 hex, 아니면 None."""
    rgb = subject_color(subject)
    if not rgb:
        return None
    inset = Emu(int(inset_cm * EMU_PER_CM))
    frame = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        inset,
        inset,
        int(slide_width) - 2 * inset,
        int(slide_height) - 2 * inset,
    )
    frame.fill.background()                       # 내부 비움(테두리만 표시)
    frame.line.color.rgb = RGBColor.from_string(rgb)
    frame.line.width = Pt(line_pt)
    frame.shadow.inherit = False
    frame.name = f"SubjectBand_{rgb}"
    return rgb
