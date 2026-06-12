from pathlib import Path

from openpyxl import load_workbook

try:
    from config.defaults import SKIP_SHEET_KEYWORDS
    from src.parse_card import is_empty_lecture, parse_lecture_sheet
    from src.validate import report_row
except ModuleNotFoundError:
    from ..config.defaults import SKIP_SHEET_KEYWORDS
    from .parse_card import is_empty_lecture, parse_lecture_sheet
    from .validate import report_row


def iter_workbooks(input_path):
    path = Path(input_path)
    if path.is_file():
        yield path
        return
    for file_path in sorted(path.glob("*.xlsx")):
        if not file_path.name.startswith("~$"):
            yield file_path


def should_skip_sheet(ws, index):
    """강좌 시트가 아닌 것만 제외한다. 시트명이 '강좌N'이 아니어도(강사가 과목·이름으로
    바꿔도) 인식되도록 prefix 요구는 없앴다. 제외 대상:
    - 첫 시트(작성 안내/표지),
    - 숨김 시트(선택목록 등 드롭다운 원본),
    - 이름에 예시/안내/선택목록 등 비강좌 키워드가 든 시트.
    빈 시트는 여기서 거르지 않고 parse 후 is_empty_lecture가 거른다."""
    if index == 0:
        return True
    if getattr(ws, "sheet_state", "visible") != "visible":
        return True
    return any(keyword in ws.title for keyword in SKIP_SHEET_KEYWORDS)


def collect_lectures(input_path):
    """입력 폴더/파일의 강좌 시트를 수집. (lectures, reports) 반환.
    열 수 없는 파일(깨짐·암호·확장자만 xlsx 등)은 그 파일만 건너뛰고 리포트에 남겨
    나머지 강사 파일 처리는 계속한다(한 파일 때문에 전체가 멈추지 않도록)."""
    lectures = []
    reports = []
    for workbook_path in iter_workbooks(input_path):
        try:
            wb = load_workbook(workbook_path, data_only=True)
        except Exception as exc:   # BadZipFile/InvalidFileException/암호 보호 등 — 파일 단위로 격리
            reports.append(
                report_row(
                    "오류",
                    {"source_file": workbook_path.name},
                    "입력 파일",
                    "FILE_READ_FAILED",
                    f"엑셀 파일을 열지 못해 건너뛰었습니다: {type(exc).__name__}",
                    raw_value=workbook_path.name,
                    suggestion="파일이 손상되었거나 암호가 걸렸는지 확인하고, .xlsx 형식으로 다시 저장해 주세요.",
                )
            )
            continue
        for index, ws in enumerate(wb.worksheets):
            if should_skip_sheet(ws, index):
                continue
            lecture = parse_lecture_sheet(ws, workbook_path.name)
            if is_empty_lecture(lecture):
                continue
            lectures.append(lecture)
    return lectures, reports
