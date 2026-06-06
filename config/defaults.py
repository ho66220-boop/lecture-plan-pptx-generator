BASE_YEAR = 2026

SKIP_SHEET_KEYWORDS = ("작성 안내", "예시", "선택목록")

FEE_TABLE = {
    "현장강의": 80000,
    "LIVE 강의": 57000,
}

FEE_OVERRIDES = {
    # "강좌ID": 320000,
}

HOLIDAY_KEYWORDS = ("휴강", "내신기간 휴강")
REVIEW_NEEDED_KEYWORDS = ("무료 영상 제공", "보강", "대체")

TEXT_LIMITS = {
    "강의 개요": 160,
    "강의특징": 450,
    "관리 프로그램": 350,
    "교재 정보": 250,
    "진도표_상세내용": 80,
}

WEEKDAYS_KR = ("월", "화", "수", "목", "금", "토", "일")

REPORT_COLUMNS = [
    "severity",
    "lecture_id",
    "source_file",
    "source_sheet",
    "field",
    "issue_code",
    "message",
    "raw_value",
    "computed_value",
    "suggestion",
]
