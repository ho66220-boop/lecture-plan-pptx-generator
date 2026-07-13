import csv
from pathlib import Path

try:
    from config.defaults import HOLIDAY_KEYWORDS, REVIEW_NEEDED_KEYWORDS
    from src.date_utils import extract_start_time_display, format_date_dot, format_period, parse_date, weekday_kr
    from src.fee import calculate_fee, stable_override_key
    from src.validate import report_row, validate_text_limits
except ModuleNotFoundError:
    from ..config.defaults import HOLIDAY_KEYWORDS, REVIEW_NEEDED_KEYWORDS
    from .date_utils import extract_start_time_display, format_date_dot, format_period, parse_date, weekday_kr
    from .fee import calculate_fee, stable_override_key
    from .validate import report_row, validate_text_limits


def load_academic_calendar(path):
    """학사일정 CSV → {날짜ISO: [event,...]} 를 (events, reports)로 반환.
    파일이 없으면 조용히 빈 dict. 읽기/디코딩 실패는 충돌 검사만 건너뛰고
    리포트에 남긴다(파이프라인 전체를 멈추지 않음 — CSV는 설정 보조 자료)."""
    if not path:
        return {}, []
    calendar_path = Path(path)
    if not calendar_path.exists():
        return {}, []

    events = {}
    try:
        with calendar_path.open("r", encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                parsed, _ = parse_date(row.get("날짜"))
                if parsed:
                    events.setdefault(parsed.isoformat(), []).append(row)
    except Exception as exc:   # 인코딩 깨짐/권한/형식 오류 등 — 검사만 건너뛰고 계속
        report = report_row(
            "경고",
            {"source_file": calendar_path.name},
            "학사일정",
            "CALENDAR_READ_FAILED",
            f"학사일정 CSV를 읽지 못해 충돌 검사를 건너뜁니다: {type(exc).__name__}",
            raw_value=str(path),
            suggestion="CSV 인코딩(UTF-8)·형식을 확인하거나 학사일정 경로를 비워 주세요.",
        )
        return {}, [report]
    return events, []


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
    # 특강 표시는 구분/시즌 어디에 적어도 인정한다. 시즌에 '특강' 단어 없이 '썸머'/'윈터'만
    # 적은 경우(예: 시즌="윈터")도 특강으로 본다 — 둘을 대칭으로 처리(과거엔 썸머만 잡혔음).
    is_special = (not is_regular) and (
        ("특강" in gubun) or ("특강" in season)
        or ("썸머" in season) or ("윈터" in season)
    )
    return is_regular, ("total" if is_special else "monthly")


def _in_target_month(parsed, year, month):
    """parsed_date가 대상 연·월에 속하는지."""
    return parsed is not None and parsed.year == year and parsed.month == month


def _date_parse_failed_report(lecture, row_idx, raw_value, excluded_month=None):
    """진도 날짜 해석 실패 리포트(월별 제외 분기/본류 공용 — 문구는 기존 그대로)."""
    if excluded_month is not None:
        message = f"진도표 {row_idx}행 날짜를 해석하지 못해 {excluded_month}월 계획서에서 제외했습니다."
        suggestion = "예: 7/20(월), 2026-07-20 형식으로 입력해 주세요."
    else:
        message = f"진도표 {row_idx}행 날짜를 해석하지 못했습니다."
        suggestion = "예: 5/4(월), 2026-05-04 형식으로 입력해 주세요."
    return report_row(
        "오류",
        lecture,
        "진도표 날짜",
        "DATE_PARSE_FAILED",
        message,
        raw_value=raw_value,
        suggestion=suggestion,
    )


def normalize_lecture(raw, index, base_year, calendar_events=None, target_month=None):
    """target_month(int, 예: 7)가 주어지면 '월별 계획서' 모드:
    정규반은 진도를 그 달(base_year+target_month)만 잘라 회차·기간·수강료를 재계산한다.
    특강(썸머특강 등)은 패키지 단위라 월로 쪼개지 않고 전체 기간 그대로 항상 포함한다.
    그 달 수업이 0회인 정규반만 (None, reports)를 돌려 슬라이드를 만들지 않는다.
    target_month=None이면 기존 동작(전체 기간, 전 강좌) 그대로."""
    fields = raw.get("fields", {})
    lecture_id = f"L{index:03d}_{raw.get('source_sheet', '')}"
    lecture = {
        "lecture_id": lecture_id,
        "source_file": raw.get("source_file", ""),
        "source_sheet": raw.get("source_sheet", ""),
        "fields": fields,
        "progress": [],
        "flags": [],
        "progress_header_found": raw.get("progress_header_found", True),
    }
    reports = []
    real_class_dates = []

    # 정규/특강 판별(청구 + 월별 필터 공유).
    is_regular, billing = classify_billing(fields)
    # 구분·시즌이 모두 비어 판별 근거가 0이면 monthly로 가정된다 — 특강인데 표기를
    # 누락하면 총액이 아닌 월 단위로 오청구될 수 있으므로 조용히 가정하지 않고 리포트.
    # (classify_billing의 반환값·계산은 불변. 회차 0이어도 발동 — 입력 완결성 신호라
    #  한 번의 실행에서 모든 수정 지점이 보이게 한다.)
    if not str(fields.get("구분") or "").strip() and not str(fields.get("시즌") or "").strip():
        reports.append(
            report_row(
                "확인필요",
                lecture,
                "구분/시즌",
                "BILLING_UNDETERMINED",
                "구분·시즌이 모두 비어 정규/특강 판별 근거가 없습니다. 월 단위(monthly) 청구로 가정했습니다.",
                suggestion="특강·썸머·윈터 강좌라면 구분 또는 시즌을 입력해 주세요(강좌 전체 합계로 청구 계산됩니다).",
            )
        )
    # 월별 모드라도 특강은 패키지 단위(썸머특강 등)라 월로 쪼개지 않고 전체 기간 그대로 넣는다.
    # → 엑셀에 있으면 어느 달 계획서를 뽑든 항상 포함. 정규반만 그 달로 슬라이싱한다.
    slice_month = target_month if (target_month is not None and is_regular) else None

    # 다음 달(미리보기용). 12월이면 다음 해 1월.
    next_year, next_month = None, None
    if slice_month is not None:
        next_month = 1 if slice_month == 12 else slice_month + 1
        next_year = base_year + 1 if slice_month == 12 else base_year

    # 진도표는 위→아래 시간순이므로, 월이 줄어들면(예: 12→1) 해를 넘긴 것으로 보고 연도를 +1.
    # (윈터 시즌처럼 12월→1월에 걸친 강좌의 날짜·요일·월 필터가 base_year 고정으로 틀어지는 것을 막음.)
    year_offset = 0
    prev_month = None

    for row_idx, row in enumerate(raw.get("progress", []), start=1):
        parsed, claimed_weekday = parse_date(row.get("날짜"), base_year=base_year)

        # 내용은 있는데 날짜만 빈 행(셀 병합으로 값이 소거된 경우 등)은 회차에서
        # 빠져 수강료가 입력 의도와 달라진다 — 조용히 넘기지 않고 리포트로 드러낸다.
        # 회차 산정 정책은 불변(날짜 없는 행 제외 유지): 리포트를 보고 사람이 입력을
        # 고쳐 재실행하는 것이 올바른 흐름. 완전 빈 행(내용도 없음)은 기존대로 무시.
        date_raw = str(row.get("날짜") or "").strip()
        if not date_raw and any(
            str(row.get(key) or "").strip() for key in ("회차", "수업 주제", "상세 내용", "비고")
        ):
            reports.append(
                report_row(
                    "오류",
                    lecture,
                    "진도표 날짜",
                    "PROGRESS_DATE_MISSING",
                    f"진도표 {row_idx}행에 내용은 있는데 날짜가 비어 있어 회차에 포함하지 못했습니다"
                    "(셀 병합 시 첫 행에만 값이 남습니다).",
                    raw_value=row_text(row),
                    suggestion="행마다 날짜를 개별 입력한 뒤 재실행해 주세요. 회차·수강료에 반영됩니다.",
                )
            )

        if parsed is not None:
            if prev_month is not None and parsed.month < prev_month:
                # 월 감소 중 진짜 연도 경계로 보는 것은 12→1(인접 wrap, 순방향 거리 1)뿐.
                # 그 외 역행(7월 진도 사이의 6월 보강행, 11→1 등)은 연도 경계인지 알 수
                # 없으므로 값을 조용히 고치지 않는다 — base_year(+기존 offset)를 유지한 채
                # 회차에 포함하고(보강행은 실제 수업) OUT_OF_ORDER_DATE로 사람이 확인.
                if (parsed.month - prev_month) % 12 == 1:
                    year_offset += 1
                    prev_month = parsed.month
                else:
                    reports.append(
                        report_row(
                            "확인필요",
                            lecture,
                            "진도표 날짜",
                            "OUT_OF_ORDER_DATE",
                            f"진도표 {row_idx}행 날짜({date_raw})가 앞 행보다 이전 달입니다. "
                            "보강·순서 어긋남이면 그대로 두어도 되지만, 연도가 바뀐 것이라면 날짜를 확인해 주세요.",
                            raw_value=date_raw,
                            computed_value=parsed.isoformat() if not year_offset else "",
                            suggestion="진도표를 날짜순으로 정렬하거나, 해를 넘긴 날짜면 연도를 명시(예: 2027-01-04)해 주세요.",
                        )
                    )
                    # prev_month는 갱신하지 않는다 — 주 흐름(직전까지의 진행 월)을 기준으로
                    # 유지해, 역행 행 하나가 이후 행들의 연도 판정을 오염시키지 않게 한다.
            else:
                prev_month = parsed.month
            if year_offset:
                try:
                    parsed = parsed.replace(year=parsed.year + year_offset)
                except ValueError:
                    # 윤년 2/29가 보정 후 평년이 되는 경우: 이 행만 날짜 해석 실패로
                    # 처리하고(아래 DATE_PARSE_FAILED 경로) 강좌·배치는 계속 진행한다.
                    parsed = None
        # 월별 모드: 이번 달 + 다음 달 행만 처리. 둘 다 아니면 제외.
        # 날짜를 못 읽는 행은 달을 알 수 없어 제외하되 리포트에 남긴다(조용히 버리지 않음).
        is_next = False
        if slice_month is not None:
            in_this = _in_target_month(parsed, base_year, slice_month)
            is_next = _in_target_month(parsed, next_year, next_month)
            if not (in_this or is_next):
                if row.get("날짜") and not parsed:
                    reports.append(
                        _date_parse_failed_report(
                            lecture, row_idx, row.get("날짜", ""), excluded_month=slice_month
                        )
                    )
                continue
        row = dict(row)
        row["parsed_date"] = parsed.isoformat() if parsed else ""
        row["computed_weekday"] = weekday_kr(parsed) if parsed else ""
        row["weekday_mismatch_flag"] = False
        row["is_real_class"] = False
        row["is_next_month"] = is_next   # 다음 달 미리보기 행(표시용 — 회차/수강료 누적 제외)

        if row.get("날짜") and not parsed:
            reports.append(_date_parse_failed_report(lecture, row_idx, row.get("날짜", "")))

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

        # 다음 달 행은 표시용 — 회차/수강료/기간 누적에는 절대 넣지 않는다(월 단가 메시지 유지).
        if parsed and not is_holiday_row(row) and not is_next:
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
    # 특강은 slice_month=None이라 이 조건을 타지 않고 전체 기간 그대로 유지된다.
    if slice_month is not None and not real_class_dates:
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
        stable_id=stable_override_key(fields.get("강사명", ""), fields.get("강좌명", "")),
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
    calendar_events, calendar_reports = load_academic_calendar(academic_calendar_path)
    normalized = []
    reports = list(calendar_reports)   # CSV 읽기 실패(CALENDAR_READ_FAILED)도 리포트에 포함
    # index는 raw 위치 기준(필터와 무관) → lecture_id가 월별 실행에서도 안정적으로 유지된다.
    for index, raw in enumerate(raw_lectures, start=1):
        # 강좌 1건의 예상 밖 예외가 배치 전체(수십 강좌)를 죽이지 않게 강좌 단위로 격리.
        # 실패한 강좌는 리포트에 "누가 왜"를 남기고 건너뛴다(다른 강좌는 정상 산출).
        try:
            lecture, lecture_reports = normalize_lecture(
                raw, index, base_year, calendar_events, target_month
            )
        except Exception as exc:
            reports.append(
                report_row(
                    "오류",
                    {
                        "lecture_id": f"L{index:03d}_{raw.get('source_sheet', '')}",
                        "source_file": raw.get("source_file", ""),
                        "source_sheet": raw.get("source_sheet", ""),
                    },
                    "정규화",
                    "LECTURE_NORMALIZE_FAILED",
                    f"강좌 정규화 중 예상 밖 오류로 이 강좌만 건너뛰었습니다: {type(exc).__name__}: {exc}",
                    suggestion="해당 시트의 입력값(날짜·필드)을 확인해 주세요. 다른 강좌는 정상 처리되었습니다.",
                )
            )
            continue
        reports.extend(lecture_reports)
        if lecture is not None:   # 월별 모드: 그 달 0회 정규반만 None → 슬라이드 제외(특강은 항상 포함)
            normalized.append(lecture)
    return normalized, reports
