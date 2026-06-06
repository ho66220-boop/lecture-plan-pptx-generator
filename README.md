# Lecture Plan PPTX Generator

강의계획서 PPTX 제작 과정에서 반복되는 수작업 입력을 줄이기 위한 개인 업무 자동화 도구입니다.
강사 입력 엑셀을 읽어 강의계획서 PPTX 초안, 정규화 데이터, 검증 리포트를 생성합니다. 완성본 생성기가 아니라 교무팀 검수 전 **초안 생성 도구**입니다.

## 목적

- 강사 입력 자료를 표준 엑셀 양식으로 정리
- 엑셀 데이터를 기반으로 강의계획서 PPTX 초안 생성
- 교무팀은 최종 검수와 미세 수정에 집중

## 사용 흐름

### Colab (권장)

1. `notebooks/강의계획서_자동화.ipynb`를 Colab에서 엽니다.
2. 셀을 순서대로 실행합니다. 입력 엑셀은 실행 중 업로드(여러 개 동시 가능)합니다.
3. 강사별 폴더로 정리된 결과 zip을 내려받아 최종 검수합니다.

### 로컬 (CLI)

```bash
pip install -r requirements.txt
python -m src.pipeline --input-path input/예시_홍길동_강의계획서.xlsx --output-dir output --base-year 2026
```

생성 결과:

- `output/generated_pptx/강의계획서_초안_YYYYMMDD_HHMM.pptx`
- `output/normalized_data.xlsx` / `.json`
- `output/validation_report.xlsx`

## 폴더 구조

```text
lecture-plan-pptx-generator/
├─ src/             자동화 로직(파이프라인·reflow·사진 등)
├─ config/          설정(수강료표·학사일정 샘플)
├─ templates/       마스터 템플릿 PPTX
├─ input/           입력 엑셀 양식 + 예시
├─ teacher_photos/  강사 사진(강사명 매칭)
├─ notebooks/       Colab 노트북
├─ design_refs/     디자인 참고
└─ tests/
```

> 이 저장소는 공개용으로 **개인정보(실제 강사 사진·실데이터)를 제외**했습니다. `teacher_photos/`와 `input/`에는 더미 예시(`홍길동`)만 포함되어 있고, 실제 사진·데이터는 `.gitignore`로 제외됩니다.

## 처리 규칙

- `강좌`로 시작하는 시트만 처리하고, 빈 강좌 시트와 `작성 안내`·`예시`·`선택목록` 계열 시트는 제외합니다.
- 진도표 날짜를 기준으로 개강일·총회차·수강기간·수강료를 자동 계산합니다.
- 날짜 파싱 실패, 요일 불일치, 학사일정 충돌, 긴 텍스트 등은 실행을 중단하지 않고 검증 리포트에 남깁니다.
- PPTX는 `templates/강의계획서_마스터템플릿.pptx`의 첫 슬라이드를 강좌 수만큼 복제한 뒤 `{{필드명}}` placeholder를 치환합니다.
- 치환 후 `src/reflow.py`가 **글자 수에 맞춰 박스를 자동 확장**합니다. 한 줄 높이를 넘치면 해당 박스와 같은 행의 배경 셀·테두리를 함께 키우고 아래 요소를 그만큼 내립니다. 가장 많이 늘어난 슬라이드에 맞춰 슬라이드 높이도 연장합니다(줄 수는 글자 폭 기반 근사 추정).
- placeholder 값 박스는 **세로 중앙 정렬**하고 상·하 여백(`PAD`)을 두어 글자가 가장자리에 붙지 않도록 합니다. 줄높이·여백 상수는 `src/reflow.py` 상단(`LINE_SLOT`, `PAD`)에서 조정합니다.
- FAQ 영역은 템플릿에서 제거되어 자동 생성물에 포함되지 않습니다(강사별로 달라 하단 여백을 수동 작성용으로 둠).
- 우상단 사진 박스에는 강사 사진을 자동으로 채웁니다(`src/teacher_photo.py`). `teacher_photos/`에 `강사명.jpg` 또는 `academy_과목_강사명.png` 형식으로 넣으면 엑셀의 `강사명`과 매칭합니다(이름 뒤 숫자·`(1)`·`[1]` 무시). 정사각 박스에 가운데 crop(cover)으로 들어가며 둥근 모서리를 유지하고, 없으면 회색 박스로 두고 `TEACHER_PHOTO_NOT_FOUND`로 표시합니다.
- 생성 후 남은 `{{...}}` placeholder는 `validation_report.xlsx`에 `UNRESOLVED_PLACEHOLDER`로 기록합니다.

## 수강료 설정

`config/defaults.py`에서 관리합니다.

- `FEE_TABLE` — 강의형태별 회차당 수강료 (현장강의 80,000원 / LIVE 강의 50,000원)
- `FEE_PER_SESSION_OVERRIDES` — **강사명 기준** 회차당 예외(강의형태 표보다 우선)
- `FEE_OVERRIDES` — 강좌ID 기준 총액 예외(필요 시)

총 수강료는 `회차당 단가 × 자동 계산된 총회차`로 산출됩니다.

## 학사일정 검증

기본 샘플은 `config/academic_calendar_sample.csv`이며, `--academic-calendar-path`로 다른 파일을 지정할 수 있습니다.
