from pathlib import Path

try:
    from config.defaults import BASE_YEAR
    from src.collect_excel import collect_lectures
    from src.export_outputs import export_normalized_data, export_validation_report
    from src.generate_pptx import generate_pptx
    from src.normalize import normalize_lectures
    from src.validate import empty_report_row
except ModuleNotFoundError:
    from ..config.defaults import BASE_YEAR
    from .collect_excel import collect_lectures
    from .export_outputs import export_normalized_data, export_validation_report
    from .generate_pptx import generate_pptx
    from .normalize import normalize_lectures
    from .validate import empty_report_row


def run_pipeline(
    input_path="input",
    output_dir="output",
    base_year=BASE_YEAR,
    academic_calendar_path=None,
    template_path=None,
    teacher_photo_dir=None,
    make_pptx=True,
    target_month=None,
):
    input_path = Path(input_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if academic_calendar_path is None:
        default_calendar = Path("config") / "academic_calendar_sample.csv"
        academic_calendar_path = default_calendar if default_calendar.exists() else None

    if teacher_photo_dir is None:
        default_photo_dir = Path("teacher_photos")
        teacher_photo_dir = str(default_photo_dir) if default_photo_dir.is_dir() else None

    raw_lectures = collect_lectures(input_path)
    # target_month(예: 7)이 주어지면 정규반 월별 계획서 모드(그 달만 잘라 재계산).
    lectures, reports = normalize_lectures(
        raw_lectures, base_year, academic_calendar_path, target_month
    )
    if not lectures:
        reports.append(empty_report_row())

    normalized_xlsx, normalized_json = export_normalized_data(lectures, output_dir)
    pptx_path = None
    if make_pptx and lectures:
        pptx_path, pptx_reports = generate_pptx(
            lectures,
            output_dir,
            template_path=template_path,
            teacher_photo_dir=teacher_photo_dir,
            target_month=target_month,
            base_year=base_year,
        )
        reports.extend(pptx_reports)
    validation_report = export_validation_report(reports, output_dir)

    return {
        "lecture_count": len(lectures),
        "normalized_xlsx": str(normalized_xlsx),
        "normalized_json": str(normalized_json),
        "validation_report": str(validation_report),
        "pptx_path": str(pptx_path) if pptx_path else "",
    }


if __name__ == "__main__":
    import argparse
    import json

    parser = argparse.ArgumentParser(description="Generate lecture-plan draft PPTX and validation outputs.")
    parser.add_argument("--input-path", default="input")
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--base-year", type=int, default=BASE_YEAR)
    parser.add_argument("--academic-calendar-path", default=None)
    parser.add_argument("--template-path", default=None)
    parser.add_argument("--teacher-photo-dir", default=None)
    parser.add_argument("--no-pptx", action="store_true")
    parser.add_argument(
        "--target-month", type=int, default=None,
        help="정규반 월별 계획서: 대상 월(예: 7). 지정 시 정규반만, 그 달 진도로 회차·기간·수강료 재계산.",
    )
    args = parser.parse_args()
    result = run_pipeline(
        input_path=args.input_path,
        output_dir=args.output_dir,
        base_year=args.base_year,
        academic_calendar_path=args.academic_calendar_path,
        template_path=args.template_path,
        teacher_photo_dir=args.teacher_photo_dir,
        make_pptx=not args.no_pptx,
        target_month=args.target_month,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
