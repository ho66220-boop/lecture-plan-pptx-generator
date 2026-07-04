import re
import unicodedata


PROGRESS_HEADERS = ("회차", "날짜", "수업 주제", "상세 내용", "비고")

# 헤더 비교는 공백을 전부 제거해서 한다("상세 내용"/"상세내용" 동일 취급).
_PROGRESS_HEADERS_NORM = tuple(h.replace(" ", "") for h in PROGRESS_HEADERS)
# 5칸 중 이만큼 이상 일치하면 진도표 헤더 행으로 인정(헤더 한 칸을 강사가 고쳐도 살림).
_HEADER_MATCH_MIN = 4


def normalize_key(value):
    text = "" if value is None else str(value).strip()
    text = re.sub(r"\s+", " ", text)
    text = text.split("(")[0].strip()
    return text


def _norm_header(value):
    """헤더 셀 비교용: normalize_key 후 공백까지 전부 제거."""
    return normalize_key(value).replace(" ", "")


def _is_progress_header_row(row):
    values = tuple(_norm_header(v) for v in row[:5])
    matched = sum(1 for got, want in zip(values, _PROGRESS_HEADERS_NORM) if got == want)
    return matched >= _HEADER_MATCH_MIN


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
        if _is_progress_header_row(row):
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

    # 파일명/시트명은 macOS 등에서 자모 분해(NFD)된 채 들어와 엑셀 셀에서 깨져 보임 → NFC로 합성.
    lecture = {
        "source_file": unicodedata.normalize("NFC", source_file or ""),
        "source_sheet": unicodedata.normalize("NFC", ws.title or ""),
        "fields": fields,
        "progress": progress_rows,
        # 진도표 헤더 행을 찾았는지. 헤더가 훼손되면 진도표가 통째로 빠지는데,
        # 그 사실이 리포트에 남도록 수집 단계(collect_excel)에서 이 플래그를 검사한다.
        "progress_header_found": progress_start is not None,
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
