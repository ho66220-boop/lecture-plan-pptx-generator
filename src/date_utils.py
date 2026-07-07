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


def parse_date(value, base_year=BASE_YEAR):
    """Return (date, claimed_weekday) from common Korean lecture-plan date text."""
    if value is None:
        return None, None
    if isinstance(value, datetime):
        return value.date(), None
    if isinstance(value, date):
        return value, None

    text = clean_text(value)
    if not text:
        return None, None

    claimed = None
    weekday_match = re.search(r"\(([월화수목금토일])\)", text)
    if weekday_match:
        claimed = weekday_match.group(1)

    # 서식이 풀려 숫자로 저장된 날짜(엑셀 일련번호)를 먼저 복구한다. 순수 숫자가 아니면 통과.
    serial_date = _excel_serial_to_date(text)
    if serial_date is not None:
        return serial_date, claimed

    for pattern in DATE_PATTERNS:
        match = pattern.search(text)
        if not match:
            continue
        year = int(match.groupdict().get("year") or base_year)
        month = int(match.group("month"))
        day = int(match.group("day"))
        try:
            return date(year, month, day), claimed
        except ValueError:
            return None, claimed
    return None, claimed


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


def extract_start_time_display(value):
    text = clean_text(value)
    if not text:
        return None
    match = re.search(r"(?P<hour>\d{1,2})\s*[:시]\s*(?P<minute>\d{2})?", text)
    if not match:
        return None
    hour = int(match.group("hour"))
    minute = int(match.group("minute") or 0)
    meridiem = "오전" if hour < 12 else "오후"
    display_hour = hour if 1 <= hour <= 12 else hour - 12
    if display_hour == 0:
        display_hour = 12
    if minute:
        return f"{meridiem} {display_hour}시 {minute}분"
    return f"{meridiem} {display_hour}시"
