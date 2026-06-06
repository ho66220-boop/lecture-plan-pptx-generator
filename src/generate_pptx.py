import copy
import re
from datetime import date, datetime
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

try:
    from src.date_utils import format_date_dot
    from src.reflow import reflow_slide
    from src.teacher_photo import apply_teacher_photo
    from src.validate import report_row
except ModuleNotFoundError:
    from .date_utils import format_date_dot
    from .reflow import reflow_slide
    from .teacher_photo import apply_teacher_photo
    from .validate import report_row


PLACEHOLDER_RE = re.compile(r"\{\{([^{}]+)\}\}")
DEFAULT_TEMPLATE_PATH = Path("templates") / "강의계획서_마스터템플릿.pptx"


def clone_template_slide(prs, template_slide):
    """Clone shapes from the template slide into a new blank slide."""
    blank_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(blank_layout)
    for shape in list(slide.shapes):
        slide.shapes._spTree.remove(shape.element)
    for shape in template_slide.shapes:
        slide.shapes._spTree.insert_element_before(copy.deepcopy(shape.element), "p:extLst")
    return slide


def remove_slide(prs, slide):
    slide_id = slide.slide_id
    slide_element = prs.slides._sldIdLst
    rel_id = None
    for item in slide_element:
        if int(item.id) == slide_id:
            rel_id = item.rId
            slide_element.remove(item)
            break
    if rel_id:
        prs.part.drop_rel(rel_id)


def text_value(value):
    if value is None:
        return ""
    return str(value)


def progress_date_display(row):
    # normalize 단계에서 base_year를 적용해 계산해 둔 parsed_date(ISO)를 그대로 사용한다.
    # (여기서 parse_date를 다시 호출하면 base_year 기본값으로 재파싱되어 요일이 틀어질 수 있음)
    iso = row.get("parsed_date")
    if iso:
        return format_date_dot(date.fromisoformat(iso))
    return text_value(row.get("날짜", ""))


def progress_content_display(row):
    values = [row.get("수업 주제", ""), row.get("상세 내용", ""), row.get("비고", "")]
    return "\n".join(text_value(value) for value in values if text_value(value).strip())


def build_placeholder_map(lecture):
    fields = lecture.get("fields", {})
    placeholder_map = {
        "강좌유형": fields.get("강좌 유형", ""),
        "구분": fields.get("구분", ""),
        "학년": fields.get("학년", ""),
        "서브슬로건": fields.get("서브 슬로건", ""),
        "메인제목": fields.get("메인 제목", ""),
        "과목": fields.get("과목", ""),
        "강사명": fields.get("강사명", ""),
        "강좌명": fields.get("강좌명", ""),
        "개강일": lecture.get("opening_date_display", ""),
        "수업시간": fields.get("수업 요일 / 시간", ""),
        "수강기간": lecture.get("period_display", ""),
        "휴강일": fields.get("휴강일 / 사유 / 수업 불가 일정", ""),
        "보강방법": fields.get("보강 방법", ""),
        "수강료": lecture.get("fee_display", ""),
        "교재정보": fields.get("교재 정보", ""),
        "강의특징": fields.get("강의특징", ""),
        "관리프로그램": fields.get("관리 프로그램", ""),
    }

    progress_rows = lecture.get("progress", [])
    for idx in range(1, 6):
        row = progress_rows[idx - 1] if len(progress_rows) >= idx else {}
        placeholder_map[f"진도_좌_{idx}_날짜"] = progress_date_display(row) if row else ""
        placeholder_map[f"진도_좌_{idx}_내용"] = progress_content_display(row) if row else ""

    for idx in range(1, 6):
        source_index = idx + 4
        row = progress_rows[source_index] if len(progress_rows) > source_index else {}
        placeholder_map[f"진도_우_{idx}_날짜"] = progress_date_display(row) if row else ""
        placeholder_map[f"진도_우_{idx}_내용"] = progress_content_display(row) if row else ""

    return {key: text_value(value) for key, value in placeholder_map.items()}


def replace_placeholders_in_text(text, placeholder_map):
    def replace(match):
        key = match.group(1).strip()
        return placeholder_map.get(key, "")

    return PLACEHOLDER_RE.sub(replace, text)


def replace_placeholders_in_paragraph(paragraph, placeholder_map):
    if not paragraph.runs:
        return
    original = "".join(run.text for run in paragraph.runs)
    if "{{" not in original:
        return
    replaced = replace_placeholders_in_text(original, placeholder_map)
    paragraph.runs[0].text = replaced
    for run in paragraph.runs[1:]:
        run.text = ""


def replace_placeholders_in_text_frame(text_frame, placeholder_map):
    for paragraph in text_frame.paragraphs:
        replace_placeholders_in_paragraph(paragraph, placeholder_map)


def iter_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)


def replace_placeholders_in_shape(shape, placeholder_map):
    if getattr(shape, "has_text_frame", False):
        replace_placeholders_in_text_frame(shape.text_frame, placeholder_map)
    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                replace_placeholders_in_text_frame(cell.text_frame, placeholder_map)


def replace_placeholders_in_slide(slide, placeholder_map):
    for shape in iter_shapes(slide.shapes):
        replace_placeholders_in_shape(shape, placeholder_map)


def collect_texts_from_slide(slide):
    texts = []
    for shape in iter_shapes(slide.shapes):
        if getattr(shape, "has_text_frame", False):
            texts.append(shape.text)
        if getattr(shape, "has_table", False):
            for row in shape.table.rows:
                for cell in row.cells:
                    texts.append(cell.text)
    return texts


def unresolved_placeholders(slide):
    found = []
    for text in collect_texts_from_slide(slide):
        found.extend(match.group(0) for match in PLACEHOLDER_RE.finditer(text or ""))
    return sorted(set(found))


def validate_required_values(lecture):
    reports = []
    fields = lecture.get("fields", {})
    required = {
        "메인 제목": fields.get("메인 제목", ""),
        "과목": fields.get("과목", ""),
        "강사명": fields.get("강사명", ""),
        "강좌명": fields.get("강좌명", ""),
    }
    for field, value in required.items():
        if not text_value(value).strip():
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


def validate_progress_overflow(lecture):
    if len(lecture.get("progress", [])) <= 10:
        return []
    return [
        report_row(
            "경고",
            lecture,
            "진도표",
            "PROGRESS_OVERFLOW",
            "진도표가 10개를 초과하여 PPTX에는 10개까지만 반영했습니다.",
            computed_value=str(len(lecture.get("progress", []))),
            suggestion="PPTX에서 추가 진도를 수동 반영하거나 템플릿 행을 확장해 주세요.",
        )
    ]


def generate_pptx_from_template(
    lectures,
    template_path,
    output_path,
    teacher_photo_dir=None,
):
    prs = Presentation(template_path)
    if not prs.slides:
        raise ValueError("Template PPTX has no slides.")

    template_slide = prs.slides[0]
    # 세로 중앙 정렬을 적용할 placeholder 값 박스 id(치환 전 템플릿 기준; 복제해도 id 유지).
    content_ids = {
        shape.shape_id
        for shape in template_slide.shapes
        if getattr(shape, "has_text_frame", False) and "{{" in shape.text_frame.text
    }
    reports = []
    generated_slides = []
    max_growth = 0

    for lecture in lectures:
        reports.extend(validate_required_values(lecture))
        reports.extend(validate_progress_overflow(lecture))
        slide = clone_template_slide(prs, template_slide)
        replace_placeholders_in_slide(slide, build_placeholder_map(lecture))
        max_growth = max(max_growth, reflow_slide(slide, content_ids))
        teacher_name = lecture.get("fields", {}).get("강사명", "")
        inserted_photo = apply_teacher_photo(slide, teacher_name, teacher_photo_dir)
        if teacher_photo_dir and not inserted_photo:
            reports.append(
                report_row(
                    "확인필요",
                    lecture,
                    "강사 사진",
                    "TEACHER_PHOTO_NOT_FOUND",
                    f"'{teacher_name}' 강사 사진을 찾지 못해 회색 박스로 둡니다.",
                    suggestion=f"{teacher_photo_dir}/{teacher_name}.jpg 형태로 파일을 넣어 주세요.",
                )
            )
        remaining = unresolved_placeholders(slide)
        for placeholder in remaining:
            reports.append(
                report_row(
                    "경고",
                    lecture,
                    "PPTX",
                    "UNRESOLVED_PLACEHOLDER",
                    f"치환되지 않은 placeholder가 남아 있습니다: {placeholder}",
                    raw_value=placeholder,
                    suggestion="placeholder_map 또는 템플릿 placeholder명을 확인해 주세요.",
                )
            )
        generated_slides.append(slide)

    remove_slide(prs, template_slide)
    # 모든 슬라이드가 공유하는 슬라이드 높이를, 가장 많이 늘어난 슬라이드에 맞춰 확장.
    if max_growth:
        prs.slide_height = int(prs.slide_height + max_growth)
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return Path(output_path), reports


def generate_pptx(lectures, output_dir, template_path=None, teacher_photo_dir=None):
    output = Path(output_dir) / "generated_pptx"
    output.mkdir(parents=True, exist_ok=True)
    template = Path(template_path or DEFAULT_TEMPLATE_PATH)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    output_path = output / f"강의계획서_초안_{timestamp}.pptx"
    return generate_pptx_from_template(
        lectures=lectures,
        template_path=template,
        output_path=output_path,
        teacher_photo_dir=teacher_photo_dir,
    )
