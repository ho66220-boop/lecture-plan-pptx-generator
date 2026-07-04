# -*- coding: utf-8 -*-
"""진도표 헤더 인식 견고성 테스트.

배경: 헤더 5칸이 정확히 일치해야만 진도표를 읽던 시절에는, 강사가 헤더를
"상세내용"처럼 살짝만 바꿔도 진도표 전체가 조용히 사라져 회차·수강료가 0으로
계산됐다(리포트에도 안 남음). 아래 테스트는 그 침묵 실패가 다시 생기지 않게 잡는다.
"""
from openpyxl import Workbook

from src.collect_excel import collect_lectures
from src.parse_card import parse_lecture_sheet


def make_sheet(header, progress_rows, fields=(("강사명", "홍길동"),)):
    wb = Workbook()
    ws = wb.active
    ws.title = "강좌1"
    for key, value in fields:
        ws.append([key, value])
    if header is not None:
        ws.append(list(header))
    for row in progress_rows:
        ws.append(list(row))
    return wb, ws


STANDARD_HEADER = ("회차", "날짜", "수업 주제", "상세 내용", "비고")
PROGRESS = [("1", "7/7(화)", "OT", "오리엔테이션", "")]


def test_standard_header_parsed():
    _, ws = make_sheet(STANDARD_HEADER, PROGRESS)
    lecture = parse_lecture_sheet(ws, "x.xlsx")
    assert lecture["progress_header_found"] is True
    assert len(lecture["progress"]) == 1


def test_header_without_spaces_still_parsed():
    # "수업주제"/"상세내용" 처럼 공백을 지워 써도 진도표가 사라지면 안 된다.
    _, ws = make_sheet(("회차", "날짜", "수업주제", "상세내용", "비고"), PROGRESS)
    lecture = parse_lecture_sheet(ws, "x.xlsx")
    assert lecture["progress_header_found"] is True
    assert len(lecture["progress"]) == 1


def test_header_with_paren_annotation_still_parsed():
    # "날짜(요일)" 같은 괄호 주석이 붙어도 인식.
    _, ws = make_sheet(("회차", "날짜(요일)", "수업 주제", "상세 내용", "비고"), PROGRESS)
    lecture = parse_lecture_sheet(ws, "x.xlsx")
    assert lecture["progress_header_found"] is True


def test_one_modified_header_cell_tolerated():
    # 5칸 중 1칸만 임의 변경("수업 내용")된 경우도 4/5 일치로 헤더 인정.
    _, ws = make_sheet(("회차", "날짜", "수업 내용", "상세 내용", "비고"), PROGRESS)
    lecture = parse_lecture_sheet(ws, "x.xlsx")
    assert lecture["progress_header_found"] is True
    assert len(lecture["progress"]) == 1


def test_missing_header_flagged_not_silent(tmp_path):
    # 헤더 행이 아예 없으면: 필드는 살고, progress_header_found=False,
    # 수집 단계에서 PROGRESS_HEADER_NOT_FOUND 오류 리포트가 남아야 한다.
    wb, ws = make_sheet(None, [])
    path = tmp_path / "깨진양식.xlsx"
    wb.save(path)

    lectures, reports = collect_lectures(path)
    assert len(lectures) == 1
    assert lectures[0]["progress_header_found"] is False
    codes = [r["issue_code"] for r in reports]
    assert "PROGRESS_HEADER_NOT_FOUND" in codes
    assert any(r["severity"] == "오류" for r in reports)
