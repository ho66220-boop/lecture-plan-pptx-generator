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


def classify_billing(fields):
    """구분/시즌으로 정규/특강을 판별. (is_regular, billing) 반환.
    정규와 특강이 동시에 붙으면(구분 혼합·여름 정규반) 정규 우선 → monthly(과다청구 방지).
    월별 계획서의 '정규반만' 필터도 이 is_regular를 그대로 공유한다."""
    gubun = fields.get("구분", "")
    season = fields.get("시즌", "")
    is_regular = "정규" in gubun
    is_special = (not is_regular) and (
        ("특강" in gubun) or ("썸머" in season) or ("특강" in season)
    )
    return is_regular, ("total" if is_special else "monthly")


def _in_target_month(parsed, year, month):
    """parsed_date가 대상 연·월에 속하는지."""
    return parsed is not None and parsed.year == year and parsed.month == month


def normalize_lecture(raw, index, base_year, calendar_events=None, target_month=None):
    """target_month(int, 예: 7)가 주어지면 '정규반 월별 계획서' 모드:
    정규반만, 진도를 그 달(base_year+target_month)만 잘라 회차·기간·수강료를 재계산한다.
    대상에서 빠지면(특강이거나 그 달 수업 0회) (None, reports)를 돌려 슬라이드를 만들지 않는다.
    target_month=None이면 기존 동작(전체 기간, 전 강좌) 그대로."""
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

    # 정규/특강 판별(청구 + 월별 필터 공유). 월별 모드에서 특강은 제외(정규반 전용).
    is_regular, billing = classify_billing(fields)
    if target_month is not None and not is_regular:
        return None, []

    for row_idx, row in enumerate(raw.get("progress", []), start=1):
        parsed, claimed_weekday = parse_date(row.get("날짜"), base_year=base_year)
        # 월별 모드: 대상 월 행만 처리. 날짜를 못 읽는 행은 달을 알 수 없어 제외하되 리포트에 남긴다.
        if target_month is not None and not _in_target_month(parsed, base_year, target_month):
            if row.get("날짜") and not parsed:
                reports.append(
                    report_row(
                        "오류",
                        lecture,
                        "진도표 날짜",
                        "DATE_PARSE_FAILED",
                        f"진도표 {row_idx}행 날짜를 해석하지 못해 {target_month}월 계획서에서 제외했습니다.",
                        raw_value=row.get("날짜", ""),
                        suggestion="예: 7/20(월), 2026-07-20 형식으로 입력해 주세요.",
                    )
                )
            continue
        row = dict(row)
        row["parsed_date"] = parsed.isoformat() if parsed else ""
        row["computed_weekday"] = weekday_kr(parsed) if parsed else ""
        row["weekday_mismatch_flag"] = False
        row["is_real_class"] = False

        if row.get("날짜") and not parsed:
            reports.append(
                report_row(
                    "오류",
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
                    "경고",
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

    # 월별 모드: 그 달에 실제 수업이 0회인 정규반은 슬라이드를 만들지 않는다(빈 장 방지).
    if target_month is not None and not real_class_dates:
        return None, reports

    first_date = min(real_class_dates) if real_class_dates else None
    last_date = max(real_class_dates) if real_class_dates else None
    total_sessions = len(real_class_dates)

    # 한 달 회차수 = 달력상 월별 회차 중 가장 많은 달(부분 월의 영향을 피하기 위해 최대값 사용).
    # 월별 모드에서는 그 달만 남아 있어 자연히 그 달 회차수가 잡힌다.
    month_counts = {}
    for class_date in real_class_dates:
        key = (class_date.year, class_date.month)
        month_counts[key] = month_counts.get(key, 0) + 1
    monthly_sessions = max(month_counts.values()) if month_counts else 0

    # 청구 단위는 위에서 판별(classify_billing) — 정규반=monthly, 특강/썸머=total.
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
        monthly_sessions=monthly_sessions,
        billing=billing,
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


def normalize_lectures(raw_lectures, base_year, academic_calendar_path=None, target_month=None):
    calendar_events = load_academic_calendar(academic_calendar_path)
    normalized = []
    reports = []
    # index는 raw 위치 기준(필터와 무관) → lecture_id가 월별 실행에서도 안정적으로 유지된다.
    for index, raw in enumerate(raw_lectures, start=1):
        lecture, lecture_reports = normalize_lecture(
            raw, index, base_year, calendar_events, target_month
        )
        reports.extend(lecture_reports)
        if lecture is not None:   # 월별 모드: 특강·그 달 0회는 None → 슬라이드 제외
            normalized.append(lecture)
    return normalized, reports
