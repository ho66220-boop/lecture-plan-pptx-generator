import csv
from pathlib import Path

try:
    from config.defaults import HOLIDAY_KEYWORDS, REVIEW_NEEDED_KEYWORDS
    from src.date_utils import extract_start_time_display, format_date_dot, format_period, parse_date, weekday_kr
    from src.fee import calculate_fee
    from src.validate import report_row, validate_text_limits
except ModuleNotFoundError:
    from ..config.defaults import HOLIDAY_KEYWORDS, REVIEW_NEEDED_KEYWORDS
    from .date_utils import extract_start_time_display, format_date_dot, format_period, parse_date, weekday_kr
    from .fee import calculate_fee
    from .validate import report_row, validate_text_limits


def load_academic_calendar(path):
    if not path:
        return {}
    calendar_path = Path(path)
    if not calendar_path.exists():
        return {}

    events = {}
    with calendar_path.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            parsed, _ = parse_date(row.get("날짜"))
            if parsed:
                events.setdefault(parsed.isoformat(), []).append(row)
    return events


def row_text(row):
    return " ".join(str(row.get(key, "") or "") for key in ("회차", "날짜", "수업 주제", "상세 내용", "비고"))


def is_holiday_row(row):
    text = row_text(row)
    return any(keyword in text for keyword in HOLIDAY_KEYWORDS)


def has_review_keyword(row):
    text = row_text(row)
    return [keyword for keyword in REVIEW_NEEDED_KEYWORDS if keyword in text]


def normalize_lecture(raw, index, base_year, calendar_events=None):
    fields = raw.get("fields", {})
    lecture_id = f"L{index:03d}_{raw.get('source_sheet', '')}"
    lecture = {
        "lecture_id": lecture_id,
        "template_type": "내신_TYPE2",
        "source_file": raw.get("source_file", ""),
        "source_sheet": raw.get("source_sheet", ""),
        "fields": fields,
        "progress": [],
        "flags": [],
    }
    reports = []
    real_class_dates = []

    for row_idx, row in enumerate(raw.get("progress", []), start=1):
        parsed, claimed_weekday = parse_date(row.get("날짜"), base_year=base_year)
        row = dict(row)
        row["parsed_date"] = parsed.isoformat() if parsed else ""
        row["computed_weekday"] = weekday_kr(parsed) if parsed else ""
        row["weekday_mismatch_flag"] = False
        row["is_real_class"] = False

        if row.get("날짜") and not parsed:
            reports.append(
                report_row(
                    "ERROR",
                    lecture,
                    "진도표 날짜",
                    "DATE_PARSE_FAILED",
                    f"진도표 {row_idx}행 날짜를 해석하지 못했습니다.",
                    raw_value=row.get("날짜", ""),
                    suggestion="예: 5/4(월), 2026-05-04 형식으로 입력해 주세요.",
                )
            )

        if parsed and claimed_weekday and claimed_weekday != row["computed_weekday"]:
            row["weekday_mismatch_flag"] = True
            lecture["flags"].append("WEEKDAY_MISMATCH")
            reports.append(
                report_row(
                    "WARNING",
                    lecture,
                    "진도표 날짜",
                    "WEEKDAY_MISMATCH",
                    f"진도표 {row_idx}행의 입력 요일과 실제 요일이 다릅니다.",
                    raw_value=row.get("날짜", ""),
                    computed_value=row["computed_weekday"],
                    suggestion="입력 요일을 확인해 주세요.",
                )
            )

        if parsed and not is_holiday_row(row):
            row["is_real_class"] = True
            real_class_dates.append(parsed)

        for keyword in has_review_keyword(row):
            lecture["flags"].append("SESSION_TYPE_REVIEW_NEEDED")
            reports.append(
                report_row(
                    "확인필요",
                    lecture,
                    "진도표",
                    "SESSION_TYPE_REVIEW_NEEDED",
                    f"진도표 {row_idx}행에 '{keyword}' 키워드가 있어 실제 회차 포함 여부 확인이 필요합니다.",
                    raw_value=row_text(row),
                    suggestion="실제 수업 회차인지 교무팀에서 확인해 주세요.",
                )
            )

        if parsed and calendar_events:
            for event in calendar_events.get(parsed.isoformat(), []):
                reports.append(
                    report_row(
                        "확인필요",
                        lecture,
                        "학사일정",
                        "ACADEMIC_CALENDAR_CONFLICT",
                        f"수업일이 학사일정 '{event.get('일정명', '')}'와 겹칩니다.",
                        raw_value=parsed.isoformat(),
                        computed_value=event.get("유형", ""),
                        suggestion=event.get("비고", "수업 진행 여부를 확인해 주세요."),
                    )
                )
                lecture["flags"].append("ACADEMIC_CALENDAR_CONFLICT")

        lecture["progress"].append(row)

    first_date = min(real_class_dates) if real_class_dates else None
    last_date = max(real_class_dates) if real_class_dates else None
    total_sessions = len(real_class_dates)

    class_time = fields.get("수업 요일 / 시간", "")
    time_display = extract_start_time_display(class_time)
    opening_display = format_date_dot(first_date)
    if opening_display and time_display:
        opening_display = f"{opening_display} {time_display}"
    elif opening_display and class_time:
        lecture["flags"].append("OPENING_TIME_REVIEW_NEEDED")
        reports.append(
            report_row(
                "확인필요",
                lecture,
                "수업 요일 / 시간",
                "OPENING_TIME_PARSE_FAILED",
                "개강일 표시에 붙일 수업 시작 시간을 자동 추출하지 못했습니다.",
                raw_value=class_time,
                computed_value=opening_display,
                suggestion="PPTX에서 수업 시간을 수동 확인해 주세요.",
            )
        )

    fee = calculate_fee(
        lecture_id,
        fields.get("강의형태", ""),
        total_sessions,
        teacher_name=fields.get("강사명", ""),
    )
    if total_sessions and not fee["fee_display"]:
        reports.append(
            report_row(
                "확인필요",
                lecture,
                "강의형태",
                "FEE_TABLE_MISSING",
                "강의형태에 맞는 회차당 수강료 설정이 없습니다.",
                raw_value=fields.get("강의형태", ""),
                suggestion="config/defaults.py의 FEE_TABLE을 확인해 주세요.",
            )
        )

    lecture.update(
        {
            "computed_opening_date": first_date.isoformat() if first_date else "",
            "opening_date_display": opening_display,
            "computed_total_sessions": total_sessions,
            "first_class_date": first_date.isoformat() if first_date else "",
            "last_class_date": last_date.isoformat() if last_date else "",
            "computed_period": format_period(first_date, last_date),
            "period_display": format_period(first_date, last_date),
            **fee,
        }
    )

    reports.extend(validate_text_limits(lecture))
    lecture["flags"] = sorted(set(lecture["flags"]))
    return lecture, reports


def normalize_lectures(raw_lectures, base_year, academic_calendar_path=None):
    calendar_events = load_academic_calendar(academic_calendar_path)
    normalized = []
    reports = []
    for index, raw in enumerate(raw_lectures, start=1):
        lecture, lecture_reports = normalize_lecture(raw, index, base_year, calendar_events)
        normalized.append(lecture)
        reports.extend(lecture_reports)
    return normalized, reports
