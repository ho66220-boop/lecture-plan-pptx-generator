try:
    from config.defaults import REPORT_COLUMNS, TEXT_LIMITS
except ModuleNotFoundError:
    from ..config.defaults import REPORT_COLUMNS, TEXT_LIMITS


def report_row(
    severity,
    lecture,
    field,
    issue_code,
    message,
    raw_value="",
    computed_value="",
    suggestion="",
):
    return {
        "severity": severity,
        "lecture_id": lecture.get("lecture_id", ""),
        "source_file": lecture.get("source_file", ""),
        "source_sheet": lecture.get("source_sheet", ""),
        "field": field,
        "issue_code": issue_code,
        "message": message,
        "raw_value": raw_value,
        "computed_value": computed_value,
        "suggestion": suggestion,
    }


def validate_text_limits(lecture):
    reports = []
    fields = lecture.get("fields", {})
    for field, limit in TEXT_LIMITS.items():
        if field.startswith("진도표_"):
            continue
        value = fields.get(field, "")
        if value and len(value) > limit:
            reports.append(
                report_row(
                    "경고",
                    lecture,
                    field,
                    "TEXT_LIMIT_EXCEEDED",
                    f"{field} 권장 글자 수({limit}자)를 초과했습니다.",
                    raw_value=value,
                    computed_value=f"{len(value)}자",
                    suggestion="PPTX에서 글자가 작아지거나 일부 생략될 수 있습니다.",
                )
            )

    detail_limit = TEXT_LIMITS.get("진도표_상세내용")
    if detail_limit:
        for idx, row in enumerate(lecture.get("progress", []), start=1):
            value = row.get("상세 내용", "")
            if value and len(value) > detail_limit:
                reports.append(
                    report_row(
                        "경고",
                        lecture,
                        "진도표_상세내용",
                        "TEXT_LIMIT_EXCEEDED",
                        f"진도표 {idx}행 상세 내용이 권장 글자 수({detail_limit}자)를 초과했습니다.",
                        raw_value=value,
                        computed_value=f"{len(value)}자",
                        suggestion="상세 내용을 줄이거나 PPTX에서 수동 조정해 주세요.",
                    )
                )
    return reports


def validate_required_values(lecture):
    """필수 필드(메인 제목·과목·강사명·강좌명) 공백과 빈 진도표 검사. 정규화 단계에서 호출되므로
    PPTX를 만들지 않는 실행(--no-pptx)에서도 리포트된다."""
    reports = []
    fields = lecture.get("fields", {})
    required = {
        "메인 제목": fields.get("메인 제목", ""),
        "과목": fields.get("과목", ""),
        "강사명": fields.get("강사명", ""),
        "강좌명": fields.get("강좌명", ""),
    }
    for field, value in required.items():
        if not str(value or "").strip():
            reports.append(
                report_row(
                    "확인필요",
                    lecture,
                    field,
                    "REQUIRED_FIELD_EMPTY",
                    f"{field} 필수값이 비어 있습니다.",
                    suggestion="강좌 입력 시트에서 값을 입력해 주세요.",
                )
            )
    if not lecture.get("progress"):
        reports.append(
            report_row(
                "확인필요",
                lecture,
                "진도표",
                "REQUIRED_PROGRESS_EMPTY",
                "진도표가 비어 있습니다.",
                suggestion="최소 1개 이상의 진도 행을 입력해 주세요.",
            )
        )
    return reports


def empty_report_row(message="처리할 강좌 시트가 없습니다."):
    return {
        column: "" for column in REPORT_COLUMNS
    } | {
        "severity": "정보",
        "issue_code": "NO_LECTURES",
        "message": message,
        "suggestion": "강좌 시트에 값을 입력한 뒤 다시 실행해 주세요.",
    }
