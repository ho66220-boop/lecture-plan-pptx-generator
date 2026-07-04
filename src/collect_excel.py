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


def should_skip_sheet(ws):
    """강좌 시트가 아닌 것만 제외한다(이름·상태 기준, 위치 무관). 제외 대상:
    - 숨김 시트(선택목록 등 드롭다운 원본),
    - 이름에 안내/예시/샘플/선택목록 등 비강좌 키워드가 든 시트(작성 안내·예시 시트).
    시트명이 '강좌N'이 아니어도(강사가 과목·이름으로 바꿔도) 인식되고, 첫 시트 위치에
    의존하지 않으므로 강사가 안내·예시 시트를 지우거나 순서를 바꿔도 강좌는 그대로 잡힌다.
    빈 시트는 여기서 거르지 않고 parse 후 is_empty_lecture가 거른다."""
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
        for ws in wb.worksheets:
            if should_skip_sheet(ws):
                continue
            lecture = parse_lecture_sheet(ws, workbook_path.name)
            if is_empty_lecture(lecture):
                continue
            # 필드는 있는데 진도표 헤더를 못 찾은 시트: 헤더가 훼손된 것.
            # 예전에는 진도표가 조용히 사라져 회차·수강료가 0으로 계산됐다 → 반드시 리포트.
            if not lecture.get("progress_header_found", True):
                reports.append(
                    report_row(
                        "오류",
                        lecture,
                        "진도표",
                        "PROGRESS_HEADER_NOT_FOUND",
                        "진도표 헤더 행(회차/날짜/수업 주제/상세 내용/비고)을 찾지 못해 진도표를 읽지 않았습니다.",
                        raw_value=lecture.get("source_sheet", ""),
                        suggestion="헤더 5칸이 양식과 같은지 확인해 주세요. 개강일·회차·수강료가 계산되지 않습니다.",
                    )
                )
            lectures.append(lecture)
    return lectures, reports
