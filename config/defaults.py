BASE_YEAR = 2026

SKIP_SHEET_KEYWORDS = ("작성 안내", "예시", "선택목록")

FEE_TABLE = {
    "현장강의": 80000,
    "LIVE 강의": 50000,
}

FEE_OVERRIDES = {
    # "강좌ID": 320000,  # 강좌ID 기준 총액 예외
}

# 강사명 기준 회차당 수강료 예외 (강의형태 표보다 우선 적용).
# 아래는 공개용 예시값입니다. 실제 운영 시 강사명·단가로 교체해 사용하세요.
FEE_PER_SESSION_OVERRIDES = {
    "홍길동": 70000,
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
