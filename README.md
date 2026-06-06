# 강의계획서 자동화

강사용 입력 엑셀을 읽어 강의계획서 PPTX 초안, 정규화 데이터, 검증 리포트를 생성하는 로컬 Python 프로젝트입니다. 완성본 생성기가 아니라 교무팀 검수 전 초안 생성 도구입니다.

## 구성

```text
lecture-plan-automation/
├─ config/
├─ src/
├─ input/
├─ design_refs/
├─ output/
├─ notebooks/
└─ tests/
```

현재 기준 입력 파일:

- `input/통합_강의계획서_입력양식.xlsx`
- `templates/강의계획서_마스터템플릿.pptx`
- `design_refs/내신_TYPE2.png`는 참고용으로 보관

## 설치

```bash
pip install -r requirements.txt
```

## 실행

```bash
python -m src.pipeline --input-path input --output-dir output --base-year 2026
```

다른 템플릿을 지정하려면:

```bash
python -m src.pipeline --input-path input --output-dir output --template-path templates/강의계획서_마스터템플릿.pptx
```

생성 결과:

- `output/generated_pptx/강의계획서_초안_YYYYMMDD_HHMM.pptx`
- `output/normalized_data.xlsx`
- `output/normalized_data.json`
- `output/validation_report.xlsx`

현재 첨부 엑셀의 `강좌1~강좌4` 시트는 비어 있으므로 그대로 실행하면 강좌 0건으로 처리됩니다. 실제 생성 테스트를 하려면 `예시_통합` 내용을 `강좌1`에 복사하거나 강좌 시트의 노란 입력 칸을 채운 뒤 실행하세요.

## 처리 규칙

- `강좌`로 시작하는 시트만 처리합니다.
- 빈 강좌 시트는 스킵합니다.
- `작성 안내`, `예시_통합`, `예시`, `선택목록` 계열 시트는 제외합니다.
- 진도표 날짜를 기준으로 개강일, 총회차, 수강기간, 수강료를 계산합니다.
- 날짜 파싱 실패, 요일 불일치, 학사일정 충돌, 긴 텍스트 등은 실행을 중단하지 않고 검증 리포트에 남깁니다.
- PPTX는 `templates/강의계획서_마스터템플릿.pptx`의 첫 번째 슬라이드를 강좌 수만큼 복제한 뒤 `{{필드명}}` placeholder만 치환합니다.
- 치환 후 `src/reflow.py`가 **글자 수에 맞춰 박스를 자동 확장**합니다. 내용이 길어 한 줄 높이를 넘치면 해당 박스와 같은 행의 배경 셀·테두리를 함께 키우고, 아래 요소를 그만큼 내립니다. 가장 많이 늘어난 슬라이드에 맞춰 슬라이드 높이도 연장합니다(줄 수는 글자 폭 기반 근사 추정).
- placeholder 값 박스는 **세로 중앙 정렬**하고 상·하 여백(`PAD`)을 두어 글자가 박스 가장자리에 붙지 않도록 합니다. 줄높이·여백 상수는 `src/reflow.py` 상단(`LINE_SLOT`, `PAD`)에서 조정합니다.
- FAQ 영역은 템플릿에서 제거되어 자동 생성물에 포함되지 않습니다. 선생님마다 FAQ가 달라 하단 여백을 수동 작성용으로 비워 둡니다. FAQ 포함 원본은 `templates/강의계획서_마스터템플릿_FAQ포함_backup.pptx`에 보관합니다.
- 우상단 사진 박스에는 강사 사진을 자동으로 채웁니다(`src/teacher_photo.py`). `teacher_photos/` 폴더에 `강사명.jpg`(엑셀의 `강사명`과 동일, jpg/png/webp 지원)로 넣어두면 매칭됩니다. 정사각 박스에 가운데 기준 crop(cover)으로 들어가며 둥근 모서리를 유지합니다. 사진이 없으면 회색 박스로 두고 검증 리포트에 `TEACHER_PHOTO_NOT_FOUND`로 표시합니다. `teacher_photo_dir`을 지정하지 않으면 `teacher_photos/` 폴더가 있을 때 자동 사용합니다.
- 생성 후 남은 `{{...}}` placeholder는 `validation_report.xlsx`에 `UNRESOLVED_PLACEHOLDER`로 기록합니다.

## 수강료 설정

`config/defaults.py`의 `FEE_TABLE`에서 강의형태별 회차당 수강료를 관리합니다.

```python
FEE_TABLE = {
    "현장강의": 80000,
    "LIVE 강의": 57000,
}
```

특정 강좌 예외 금액은 `FEE_OVERRIDES`에 강좌ID 기준으로 추가할 수 있습니다.

## 학사일정 검증

기본 샘플 파일은 `config/academic_calendar_sample.csv`입니다.

```csv
날짜,일정명,유형,비고
2026-06-04,더프 모의고사,더프,오후 수업 주의
```

다른 파일을 쓰려면 다음처럼 지정합니다.

```bash
python -m src.pipeline --academic-calendar-path config/academic_calendar_sample.csv
```
