# -*- coding: utf-8 -*-
"""골든 파일 E2E: 예시 엑셀을 파이프라인에 통째로 넣고 산출값을 스냅샷으로 고정.

순수 함수 단위 테스트가 못 잡는 것 — 엑셀 파싱, 정규화, PPTX 생성, 리포트 내보내기의
연결 전체 — 을 이 한 개가 잡는다. 예시 엑셀이나 계산 규칙을 의도적으로 바꿨다면
아래 기대값도 함께 갱신할 것.
"""
import json
from pathlib import Path

import pytest

from src.pipeline import run_pipeline

SAMPLE = Path("input") / "예시_홍길동_강의계획서.xlsx"


@pytest.fixture(scope="module")
def result(tmp_path_factory):
    output_dir = tmp_path_factory.mktemp("e2e_out")
    return run_pipeline(
        input_path=str(SAMPLE),
        output_dir=str(output_dir),
        base_year=2026,
    ), output_dir


def test_outputs_exist(result):
    res, _ = result
    assert res["lecture_count"] == 1
    assert Path(res["normalized_xlsx"]).exists()
    assert Path(res["normalized_json"]).exists()
    assert Path(res["validation_report"]).exists()
    assert res["pptx_path"] and Path(res["pptx_path"]).exists()


def test_report_counts_present(result):
    res, _ = result
    counts = res["report_counts"]
    assert set(counts) == {"오류", "경고", "확인필요", "정보"}
    # 예시 엑셀에는 요일 불일치 1건이 심어져 있다(경고). 오류는 없어야 정상.
    assert counts["오류"] == 0
    assert counts["경고"] >= 1


def test_normalized_snapshot(result):
    res, _ = result
    data = json.loads(Path(res["normalized_json"]).read_text(encoding="utf-8"))
    lecture = data[0] if isinstance(data, list) else data

    # 진도표 → 자동 계산 핵심값 스냅샷. 여기가 틀리면 수강료·개강일이 틀린 것.
    assert lecture["computed_total_sessions"] == 6
    assert lecture["first_class_date"] == "2026-07-07"
    assert lecture["last_class_date"] == "2026-08-11"
    assert lecture["opening_date_display"] == "07. 07(화) 오후 6시 30분"
    assert lecture["computed_period"] == "7/7(화) ~ 8/11(화)"

    # 홍길동은 강사별 회차당 예외(70,000원) + 정규반 월 단위 청구.
    assert lecture["billing"] == "monthly"
    assert lecture["fee_per_session"] == 70000
    assert lecture["computed_fee"] == 280000
    assert lecture["fee_display"] == "월 280,000원 (회차당 70,000원)"

    # 예시에 심어둔 요일 불일치 플래그가 살아 있어야 한다.
    assert "WEEKDAY_MISMATCH" in lecture["flags"]
    assert lecture["progress_header_found"] is True
