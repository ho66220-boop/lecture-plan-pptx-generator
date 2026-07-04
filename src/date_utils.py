import re
from datetime import date, datetime

try:
    from config.defaults import BASE_YEAR, WEEKDAYS_KR
except ModuleNotFoundError:
    from ..config.defaults import BASE_YEAR, WEEKDAYS_KR


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
