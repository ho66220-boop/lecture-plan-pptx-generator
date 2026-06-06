import re


PROGRESS_HEADERS = ("회차", "날짜", "수업 주제", "상세 내용", "비고")


def normalize_key(value):
    text = "" if value is None else str(value).strip()
    text = re.sub(r"\s+", " ", text)
    text = text.split("(")[0].strip()
    return text


def stringify(value):
    if value is None:
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def first_value(row, start=1, end=5):
    for cell in row[start:end]:
        value = stringify(cell)
        if value:
            return value
    return ""


def parse_lecture_sheet(ws, source_file):
    fields = {}
    progress_rows = []
    progress_start = None

    rows = list(ws.iter_rows(values_only=True))
    for idx, row in enumerate(rows, start=1):
        label = normalize_key(row[0] if row else "")
        if not label:
            continue
        header_values = tuple(normalize_key(v) for v in row[:5])
        if header_values == PROGRESS_HEADERS:
            progress_start = idx + 1
            break
        if label.startswith("[") or "자동 산출" in label:
            continue
        value = first_value(row)
        if value:
            fields[label] = value

    if progress_start:
        for row in rows[progress_start - 1 :]:
            values = [stringify(v) for v in row[:5]]
            if not any(values):
                continue
            progress_rows.append(dict(zip(PROGRESS_HEADERS, values)))

    lecture = {
        "source_file": source_file,
        "source_sheet": ws.title,
        "fields": fields,
        "progress": progress_rows,
    }
    return lecture


def is_empty_lecture(lecture):
    meaningful_fields = [
        value for key, value in lecture["fields"].items() if key not in ("통합 강의계획서 입력 양식 v2",)
    ]
    meaningful_progress = [
        row for row in lecture["progress"] if any(row.get(k) for k in PROGRESS_HEADERS)
    ]
    return not any(meaningful_fields) and not any(meaningful_progress)
