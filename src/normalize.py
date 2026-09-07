import csv
import re
from pathlib import Path

try:
    from config.defaults import BASE_YEAR, HOLIDAY_KEYWORDS, HOLIDAY_NEGATION_WORDS, REVIEW_NEEDED_KEYWORDS
    from src.date_utils import (
        extract_dates, extract_start_time, format_date_dot, format_period, parse_date_detail, weekday_kr,
    )
    from src.fee import calculate_fee, stable_override_key, unused_override_keys
    from src.validate import report_row, validate_required_values, validate_text_limits
except ModuleNotFoundError:
    from ..config.defaults import BASE_YEAR, HOLIDAY_KEYWORDS, HOLIDAY_NEGATION_WORDS, REVIEW_NEEDED_KEYWORDS
    from .date_utils import (
        extract_dates, extract_start_time, format_date_dot, format_period, parse_date_detail, weekday_kr,
    )
    from .fee import calculate_fee, stable_override_key, unused_override_keys
    from .validate import report_row, validate_required_values, validate_text_limits


HOLIDAY_FIELD = "휴강일 / 사유 / 수업 불가 일정"


def load_academic_calendar(path, base_year=BASE_YEAR):
    """학사일정 CSV → {날짜ISO: [event,...]} 를 (events, reports)로 반환.
    읽기/디코딩 실패는 충돌 검사만 건너뛰고 리포트에 남긴다(파이프라인 전체를 멈추지 않음 —
    CSV는 설정 보조 자료). 경로를 지정했는데 파일이 없거나, 날짜를 못 읽은 행이 있거나,
    결과적으로 일정이 0건이라 검사가 사실상 꺼진 상태도 모두 리포트한다 — '검사가 돌았는지'를
    리포트만 보고 알 수 있어야 하기 때문(이전에는 전부 조용히 넘어갔다).
    연도 없는 날짜(7/7)는 실행의 base_year로 해석한다."""
    if not path:
        return {}, []
    calendar_path = Path(path)
    calendar_ref = {"source_file": calendar_path.name}
    if not calendar_path.exists():
        return {}, [
            report_row(
                "경고",
                calendar_ref,
                "학사일정",
                "CALENDAR_NOT_FOUND",
                "지정한 학사일정 CSV 파일이 없어 충돌 검사를 건너뜁니다.",
                raw_value=str(path),
                suggestion="경로를 확인하거나 학사일정 경로를 비워 주세요.",
            )
        ]

    events = {}
    reports = []
    try:
        with calendar_path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            header = [(h or "").strip() for h in (reader.fieldnames or [])]
            if "날짜" not in header:
                return {}, [
                    report_row(
                        "경고",
                        calendar_ref,
                        "학사일정",
                        "CALENDAR_HEADER_MISSING",
                        "학사일정 CSV에 '날짜' 헤더 열이 없어 충돌 검사를 건너뜁니다.",
                        raw_value=",".join(header),
                        suggestion="첫 행을 '날짜,일정명,유형,비고' 형식으로 맞춰 주세요.",
                    )
                ]
            for line_no, row in enumerate(reader, start=2):
                row = {(k or "").strip(): v for k, v in row.items()}
                date_text = str(row.get("날짜") or "").strip()
                if not date_text or date_text.startswith("#"):
                    continue   # 빈 행·주석 행
                parsed, _, _ = parse_date_detail(date_text, base_year=base_year)
                if parsed:
                    events.setdefault(parsed.isoformat(), []).append(row)
                else:
                    reports.append(
                        report_row(
                            "경고",
                            calendar_ref,
                            "학사일정",
                            "CALENDAR_ROW_SKIPPED",
                            f"학사일정 CSV {line_no}행의 날짜를 해석하지 못해 그 일정은 검사에서 빠졌습니다.",
                            raw_value=date_text,
                            suggestion="예: 7/20, 2026-07-20 형식으로 입력해 주세요.",
                        )
                    )
    except Exception as exc:   # 인코딩 깨짐/권한/형식 오류 등 — 검사만 건너뛰고 계속
        report = report_row(
            "경고",
            calendar_ref,
            "학사일정",
            "CALENDAR_READ_FAILED",
            f"학사일정 CSV를 읽지 못해 충돌 검사를 건너뜁니다: {type(exc).__name__}",
            raw_value=str(path),
            suggestion="CSV 인코딩(UTF-8)·형식을 확인하거나 학사일정 경로를 비워 주세요.",
        )
        return {}, [report]
    if not events:
        reports.append(
            report_row(
                "정보",
                calendar_ref,
                "학사일정",
                "CALENDAR_EMPTY",
                "학사일정 CSV에 일정이 0건이라 학사일정 충돌 검사는 수행되지 않았습니다.",
                raw_value=str(path),
                suggestion="휴원일·시험 기간 등을 CSV에 입력하면 수업일과의 충돌을 자동 검사합니다.",
            )
        )
    return events, reports


def row_text(row):
    return " ".join(str(row.get(key, "") or "") for key in ("회차", "날짜", "수업 주제", "상세 내용", "비고"))


_NEGATION_ALT = "|".join(re.escape(w) for w in sorted(HOLIDAY_NEGATION_WORDS, key=len, reverse=True))
_HOLIDAY_NEGATION_RE = re.compile(
    "(?:" + "|".join(re.escape(k) for k in HOLIDAY_KEYWORDS) + r")\s*[:：]?\s*(?:" + _NEGATION_ALT + r")(?![가-힣])"
)


def holiday_text(text):
    """휴강 부정 표현("휴강 없음"·"휴강X")을 걷어낸 뒤 휴강 키워드가 남아 있는지."""
    stripped = _HOLIDAY_NEGATION_RE.sub(" ", text or "")
    return any(keyword in stripped for keyword in HOLIDAY_KEYWORDS)


def is_holiday_row(row):
    return holiday_text(row_text(row))


def has_review_keyword(row):
    text = row_text(row)
    return [keyword for keyword in REVIEW_NEEDED_KEYWORDS if keyword in text]


def classify_billing(fields):
    """구분/시즌으로 정규/특강을 판별. (is_regular, billing) 반환.
    정규와 특강이 동시에 붙으면(구분 혼합·여름 정규반) 정규 우선 → monthly(과다청구 방지).
    월별 계획서의 '정규반만' 필터도 이 is_regular를 그대로 공유한다."""
    gubun = fields.get("구분", "")
    season = fields.get("시즌", "")
    # 정규 표기도 구분/시즌 어디에 적어도 인정한다 — 시즌="정규반", 구분="단과"처럼 적으면
    # 월 단위로 청구하면서 월별 필터는 타지 않는 불일치가 있었다(청구 판정과 필터 판정 공유).
    is_regular = ("정규" in gubun) or ("정규" in season)
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
    # 휴강일 칸(자유 텍스트)에 적힌 날짜 — 진도표와 대조해 서로 어긋나면 리포트(아래 참조).
    holiday_field_text = str(fields.get(HOLIDAY_FIELD) or "")
    holiday_field_dates = set(extract_dates(holiday_field_text, base_year=base_year))
    holiday_rows = []   # (row_idx, parsed) — 진도표에서 휴강으로 처리한 행

    for row_idx, row in enumerate(raw.get("progress", []), start=1):
        parsed, claimed_weekday, year_explicit = parse_date_detail(row.get("날짜"), base_year=base_year)

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
            # 연도가 명시된 날짜(2027-01-04·엑셀 날짜 셀)는 이미 맞는 연도라 보정하지 않는다
            # (이전에는 12→1 경계 뒤의 명시 연도까지 +1되어 2027→2028로 틀어졌다).
            if year_offset and not year_explicit:
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
        holiday = is_holiday_row(row)
        if parsed and not holiday and not is_next:
            row["is_real_class"] = True
            real_class_dates.append(parsed)
        if parsed and holiday:
            holiday_rows.append((row_idx, parsed))

        # 휴강일 칸에 적힌 날짜가 진도표에서는 정상 수업으로 계산되는 경우 — 회차·수강료가
        # 입력 의도와 달라질 수 있다. 값은 고치지 않고(정책: 애매하면 리포트) 사람이 확인.
        if parsed and not holiday and not is_next and parsed in holiday_field_dates:
            lecture["flags"].append("HOLIDAY_FIELD_CONFLICT")
            reports.append(
                report_row(
                    "확인필요",
                    lecture,
                    "휴강일 / 진도표",
                    "HOLIDAY_FIELD_CONFLICT",
                    f"휴강일 칸에 적힌 {parsed.month}/{parsed.day}이(가) 진도표 {row_idx}행에서는 "
                    "정상 수업으로 계산됐습니다.",
                    raw_value=holiday_field_text,
                    computed_value=row_text(row),
                    suggestion="휴강이 맞으면 진도표 해당 행 비고에 '휴강'을 적어 주세요(회차·수강료에서 제외됩니다).",
                )
            )

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

    # 반대 방향 대조: 휴강일 칸에 날짜를 적었는데 진도표의 휴강 행이 그 목록에 없는 경우.
    # (휴강일 칸에 날짜가 하나도 없으면 — "내신기간 휴강"처럼 서술만 있으면 — 대조하지 않는다.)
    if holiday_field_dates:
        for row_idx, parsed in holiday_rows:
            if parsed not in holiday_field_dates:
                lecture["flags"].append("HOLIDAY_FIELD_CONFLICT")
                reports.append(
                    report_row(
                        "확인필요",
                        lecture,
                        "휴강일 / 진도표",
                        "HOLIDAY_FIELD_CONFLICT",
                        f"진도표 {row_idx}행({parsed.month}/{parsed.day})은 휴강으로 처리됐지만 휴강일 칸에는 "
                        "그 날짜가 없습니다.",
                        raw_value=holiday_field_text,
                        computed_value=parsed.isoformat(),
                        suggestion="휴강일 칸과 진도표 중 어느 쪽이 맞는지 확인해 주세요.",
                    )
                )

    # 필수 입력값·빈 진도표 검사는 산출 형식(PPTX 생성 여부)과 무관하게 항상 수행한다
    # (이전에는 PPTX 단계에만 있어 --no-pptx 검증 실행에서 빠졌다).
    reports.extend(validate_required_values(lecture))

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
    time_info = extract_start_time(class_time)
    time_display = time_info["display"] if time_info else None
    opening_display = format_date_dot(first_date)
    if opening_display and time_display:
        opening_display = f"{opening_display} {time_display}"
        if time_info["ambiguous"]:
            # "화 2:00~5:00"처럼 오전/오후 없이 1~6시로 적힌 경우 — 오후일 가능성이 높지만
            # 단정하지 않고 표시는 입력 그대로(오전) 두고 확인을 요청한다.
            lecture["flags"].append("OPENING_TIME_REVIEW_NEEDED")
            reports.append(
                report_row(
                    "확인필요",
                    lecture,
                    "수업 요일 / 시간",
                    "OPENING_TIME_AMBIGUOUS",
                    f"수업 시작 시간에 오전/오후 표기가 없어 '{time_display}'로 표시했습니다. 오후 수업이면 잘못된 표기입니다.",
                    raw_value=class_time,
                    computed_value=time_display,
                    suggestion="수업 요일 / 시간에 '오후 2:00' 또는 '14:00'처럼 오전/오후를 분명히 적어 주세요.",
                )
            )
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


def validate_target_month(target_month):
    """대상 월은 None 또는 1~12. 13 같은 값은 정규반이 전부 조용히 제외되므로 즉시 실패."""
    if target_month is None:
        return None
    if isinstance(target_month, bool) or not isinstance(target_month, int) or not 1 <= target_month <= 12:
        raise ValueError(f"target_month(대상 월)는 1~12 사이 정수여야 합니다: {target_month!r}")
    return target_month


def fee_override_reports(raw_lectures):
    """어떤 강좌와도 매칭되지 않은 FEE_OVERRIDES 키를 정보 리포트로. 필터(월별 모드) 이전의
    전체 입력 강좌를 기준으로 하므로 그 달 수업이 없어 제외된 강좌의 예외는 미사용으로 보지 않는다."""
    stable_ids = []
    lecture_ids = []
    for index, raw in enumerate(raw_lectures, start=1):
        fields = raw.get("fields", {})
        stable_ids.append(stable_override_key(fields.get("강사명", ""), fields.get("강좌명", "")))
        lecture_ids.append(f"L{index:03d}_{raw.get('source_sheet', '')}")
    return [
        report_row(
            "정보",
            {"source_file": "config/defaults.py"},
            "FEE_OVERRIDES",
            "FEE_OVERRIDE_UNUSED",
            f"수강료 총액 예외 '{key}'가 이번 입력의 어떤 강좌와도 맞지 않아 적용되지 않았습니다.",
            raw_value=key,
            suggestion="강사명·강좌명 오타나 강좌명 변경 여부를 확인해 주세요. 이번 입력에 없는 강좌면 무시해도 됩니다.",
        )
        for key in unused_override_keys(stable_ids, lecture_ids)
    ]


def normalize_lectures(raw_lectures, base_year, academic_calendar_path=None, target_month=None):
    target_month = validate_target_month(target_month)
    calendar_events, calendar_reports = load_academic_calendar(academic_calendar_path, base_year=base_year)
    normalized = []
    reports = list(calendar_reports)   # 학사일정 읽기·건너뜀·0건 리포트도 포함
    reports.extend(fee_override_reports(raw_lectures))
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
