BASE_YEAR = 2026

SKIP_SHEET_KEYWORDS = ("안내", "예시", "샘플", "선택목록")

FEE_TABLE = {
    "현장강의": 80000,
    "LIVE 강의": 50000,
}

FEE_OVERRIDES = {
    # 총액 예외. 키는 "강사명|강좌명"(공백·띄어쓰기 무시) — 파일 추가·순서 변경·시트명 변경에도 안 바뀌는 안정 키.
    # 예: "홍길동|A고 국어 내신 대비반": 320000,
    # 어떤 강좌와도 매칭되지 않은 키는 실행 시 FEE_OVERRIDE_UNUSED(정보)로 리포트된다 — 오타·강좌명 변경 확인용.
    # (구버전 "L001_강좌1" 같은 lecture_id 키도 아직 조회되지만 순서가 밀리면 다른 강좌에 붙으니 이관 권장.)
}

# 강사명 기준 회차당 수강료 예외 (강의형태 표보다 우선 적용).
# 아래는 공개용 예시값입니다. 실제 운영 시 강사명·단가로 교체해 사용하세요.
FEE_PER_SESSION_OVERRIDES = {
    "홍길동": 70000,
}

HOLIDAY_KEYWORDS = ("휴강", "내신기간 휴강")
# 휴강 키워드 뒤에 이런 말이 붙으면 '휴강이 아님'으로 본다(예: "휴강 없음", "휴강 아님", "휴강X").
# 이전에는 단순 부분 일치라 "휴강 없음" 행도 회차에서 빠져 수강료가 줄었다.
HOLIDAY_NEGATION_WORDS = ("없음", "없슴", "없다", "없이", "아님", "아니", "안함", "안 함", "무", "X", "x", "×")
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
