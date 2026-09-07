import re
from datetime import date, datetime, timedelta

try:
    from config.defaults import BASE_YEAR, WEEKDAYS_KR
except ModuleNotFoundError:
    from ..config.defaults import BASE_YEAR, WEEKDAYS_KR


# 엑셀 날짜 일련번호 처리: 날짜 서식이 풀린 셀은 openpyxl이 '46225' 같은 숫자(1900 기준 일련번호)로
# 돌려줘 날짜로 안 잡힌다. 오변환을 막기 위해 (1) '날짜 칸의 순수 숫자'이고 (2) 현대 강의 날짜 범위에
# 드는 값만 변환한다. 회차 수("3")·연도("2026")·구분자 없는 오타("20260706")는 범위 밖이라 안 건드린다.
EXCEL_DATE_EPOCH = date(1899, 12, 30)   # 1900 날짜 시스템(1900 윤년 버그 포함) 기준점
EXCEL_SERIAL_MIN = 40000                # ≈ 2009-07-06
EXCEL_SERIAL_MAX = 60000                # ≈ 2064-04-08


def _excel_serial_to_date(text):
    """순수 숫자 문자열이 엑셀 날짜 일련번호(현대 범위)면 date로, 아니면 None."""
    try:
        serial = int(float(text))   # '46225' / '46225.0' / 시간분수 '46225.5' 모두 그날로
    except (TypeError, ValueError):
        return None                 # '7/6' 등 숫자 아닌 값은 float에서 실패 → 변환 안 함
    if EXCEL_SERIAL_MIN <= serial <= EXCEL_SERIAL_MAX:
        try:
            return EXCEL_DATE_EPOCH + timedelta(days=serial)
        except (OverflowError, OSError):
            return None
    return None


DATE_PATTERNS = (
    re.compile(r"(?P<year>20\d{2})\s*[-./]\s*(?P<month>\d{1,2})\s*[-./]\s*(?P<day>\d{1,2})"),
    re.compile(r"(?P<month>\d{1,2})\s*월\s*(?P<day>\d{1,2})\s*일?"),
    # "3-4교시", "1-2회" 같은 범위 표기를 3월 4일로 오인하지 않도록 뒤에 단위어가 오면 제외.
    # (?!\d): 단위어 회피용 백트래킹으로 day가 잘리는 것 방지("3/15 회의"가 3/1로 왜곡되는 회귀 차단).
    # 회(?![가-힣]): '1-2회'는 단위로 보되 '회의실·회식'처럼 단어의 첫 글자인 '회'는 단위가 아님.
    re.compile(r"(?P<month>\d{1,2})\s*[/.-]\s*(?P<day>\d{1,2})(?!\d)(?!\s*(?:교시|회차|차시|회(?![가-힣])))"),
)


def clean_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def weekday_kr(value):
    return WEEKDAYS_KR[value.weekday()]


def parse_date_detail(value, base_year=BASE_YEAR):
    """(date, claimed_weekday, year_explicit) 반환.
    year_explicit=True는 입력 자체에 연도가 있었다는 뜻(ISO 표기·엑셀 날짜 셀·일련번호).
    연도가 명시된 날짜는 12→1월 경계의 연도 추론(normalize의 year_offset) 대상이 아니다 —
    이미 맞는 연도를 한 번 더 올려 2027→2028로 바뀌는 회귀를 막는다."""
    if value is None:
        return None, None, False
    if isinstance(value, datetime):
        return value.date(), None, True
    if isinstance(value, date):
        return value, None, True

    text = clean_text(value)
    if not text:
        return None, None, False

    claimed = None
    weekday_match = re.search(r"\(([월화수목금토일])\)", text)
    if weekday_match:
        claimed = weekday_match.group(1)

    # 서식이 풀려 숫자로 저장된 날짜(엑셀 일련번호)를 먼저 복구한다. 순수 숫자가 아니면 통과.
    serial_date = _excel_serial_to_date(text)
    if serial_date is not None:
        return serial_date, claimed, True

    for pattern in DATE_PATTERNS:
        match = pattern.search(text)
        if not match:
            continue
        explicit_year = match.groupdict().get("year")
        year = int(explicit_year or base_year)
        month = int(match.group("month"))
        day = int(match.group("day"))
        try:
            return date(year, month, day), claimed, bool(explicit_year)
        except ValueError:
            return None, claimed, bool(explicit_year)
    return None, claimed, False


def parse_date(value, base_year=BASE_YEAR):
    """Return (date, claimed_weekday) from common Korean lecture-plan date text."""
    parsed, claimed, _ = parse_date_detail(value, base_year=base_year)
    return parsed, claimed


def extract_dates(value, base_year=BASE_YEAR):
    """자유 텍스트(예: 휴강일 칸 "7/14, 7/21 내신 휴강")에서 날짜를 전부 뽑아 date 목록으로.
    연도 표기(2026-07-14)를 먼저 걷어낸 뒤 월/일 표기를 찾아, '2026-07-14'의 '26-07'이
    월/일로 오인되는 것을 막는다. 해석 불가한 조각은 조용히 건너뛴다(보조 검증용)."""
    text = clean_text(value)
    if not text:
        return []
    found = []

    def _add(year, month, day):
        try:
            found.append(date(int(year), int(month), int(day)))
        except ValueError:
            pass

    iso_pattern = DATE_PATTERNS[0]
    for match in iso_pattern.finditer(text):
        _add(match.group("year"), match.group("month"), match.group("day"))
    text = iso_pattern.sub(" ", text)
    for pattern in DATE_PATTERNS[1:]:
        for match in pattern.finditer(text):
            _add(base_year, match.group("month"), match.group("day"))
        text = pattern.sub(" ", text)
    return found


def format_date_dot(value):
    if not value:
        return ""
    return f"{value.month:02d}. {value.day:02d}({weekday_kr(value)})"


def format_period(first_date, last_date):
    if not first_date or not last_date:
        return ""
    if first_date == last_date:
        return format_date_dot(first_date)
    return f"{first_date.month}/{first_date.day}({weekday_kr(first_date)}) ~ {last_date.month}/{last_date.day}({weekday_kr(last_date)})"


# 시작 시간 해석: 오전/오후(한글·영문·구어 표현)를 반영하고 시·분 범위를 검증한다.
# 이전에는 숫자만 뽑아 "오후 2:00"→"오전 2시", "25:99"→"오후 13시 99분"으로 생성됐다.
_PM_WORDS = ("오후", "저녁", "밤", "pm", "p.m.", "p.m")
_AM_WORDS = ("오전", "새벽", "아침", "am", "a.m.", "a.m")
_NOON_WORDS = ("낮", "정오")
_MERIDIEM_RE = r"(?P<mer>오전|오후|저녁|새벽|아침|낮|정오|밤|[AaPp]\.?[Mm]\.?)"
_TIME_RE = re.compile(
    rf"(?:{_MERIDIEM_RE}\s*)?"
    r"(?P<hour>\d{1,2})\s*(?:[:：]\s*(?P<minute>\d{2})|시\s*(?:(?P<minute_kr>\d{1,2})\s*분)?)"
    r"(?:\s*(?P<mer_after>[AaPp]\.?[Mm]\.?))?"
)
# 오전/오후 표기 없이 1~6시로 적힌 경우: 학원 수업은 보통 오후지만 단정할 수 없어 표시는
# 24시간제 그대로(오전) 두고 확인필요로 알린다.
AMBIGUOUS_HOUR_MAX = 6


def extract_start_time(value):
    """수업 시간 텍스트에서 시작 시간을 해석해 dict 반환(해석 불가면 None).
    {"display": "오후 6시 30분", "hour": 18, "minute": 30, "ambiguous": bool}"""
    text = clean_text(value)
    if not text:
        return None
    match = _TIME_RE.search(text)
    if not match:
        return None
    hour = int(match.group("hour"))
    minute = int(match.group("minute") or match.group("minute_kr") or 0)
    meridiem = (match.group("mer") or match.group("mer_after") or "").lower()
    if hour > 23 or minute > 59:
        return None
    ambiguous = False
    if meridiem:
        if hour > 12:
            pass   # "오후 14:00"처럼 24시간제와 섞어 쓴 경우 시각 그대로 신뢰
        elif meridiem in _PM_WORDS or meridiem.startswith("p"):
            if hour < 12:
                hour += 12
        elif meridiem in _AM_WORDS or meridiem.startswith("a"):
            if hour == 12:
                hour = 0
        elif meridiem in _NOON_WORDS:
            if hour < 6:   # "낮 1시" → 13시
                hour += 12
    elif 1 <= hour <= AMBIGUOUS_HOUR_MAX:
        ambiguous = True
    display_meridiem = "오전" if hour < 12 else "오후"
    display_hour = hour % 12 or 12
    if minute:
        display = f"{display_meridiem} {display_hour}시 {minute}분"
    else:
        display = f"{display_meridiem} {display_hour}시"
    return {"display": display, "hour": hour, "minute": minute, "ambiguous": ambiguous}


def extract_start_time_display(value):
    info = extract_start_time(value)
    return info["display"] if info else None
