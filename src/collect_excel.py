from pathlib import Path

from openpyxl import load_workbook

try:
    from config.defaults import SKIP_SHEET_KEYWORDS
    from src.parse_card import is_empty_lecture, parse_lecture_sheet
except ModuleNotFoundError:
    from ..config.defaults import SKIP_SHEET_KEYWORDS
    from .parse_card import is_empty_lecture, parse_lecture_sheet


def iter_workbooks(input_path):
    path = Path(input_path)
    if path.is_file():
        yield path
        return
    for file_path in sorted(path.glob("*.xlsx")):
        if not file_path.name.startswith("~$"):
            yield file_path


def should_skip_sheet(sheet_name):
    if not sheet_name.startswith("강좌"):
        return True
    return any(keyword in sheet_name for keyword in SKIP_SHEET_KEYWORDS)


def collect_lectures(input_path):
    lectures = []
    for workbook_path in iter_workbooks(input_path):
        wb = load_workbook(workbook_path, data_only=True)
        for ws in wb.worksheets:
            if should_skip_sheet(ws.title):
                continue
            lecture = parse_lecture_sheet(ws, workbook_path.name)
            if is_empty_lecture(lecture):
                continue
            lectures.append(lecture)
    return lectures
