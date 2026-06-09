# Colab 실행 셀

## 1. 패키지 설치

```python
!pip install pandas openpyxl python-pptx
```

## 2. 프로젝트 준비

프로젝트 폴더를 Colab에 업로드하거나 Google Drive를 마운트합니다.

```python
from google.colab import drive
drive.mount("/content/drive")
```

예시 경로:

```python
%cd /content/lecture-plan-automation
```

또는 Drive에 둔 경우:

```python
%cd /content/drive/MyDrive/lecture-plan-automation
```

## 3. 입력 파일 배치

- 강사용 입력 엑셀: `input/통합_강의계획서_입력양식.xlsx` (또는 강사가 보낸 파일을 `--input-path`로 지정)
- 마스터 템플릿: `templates/강의계획서_마스터템플릿.pptx`
- 디자인 참고 이미지: `design_refs/내신_TYPE2.png`는 참고용

> **중요**: `templates/강의계획서_마스터템플릿.pptx`에는 이미 **FAQ가 제거**돼 있고, 박스 자동 확장·세로 중앙 정렬 로직은 `run_pipeline` 실행 시 코드(`src/reflow.py`)에서 자동 적용됩니다. FAQ 포함 원본(`_FAQ포함_backup.pptx`)으로 덮어쓰지 마세요. 폴더 전체를 그대로 업로드하면 윈도우/콜랩 어디서 돌려도 **동일한 PPTX**가 나옵니다(줄 수 추정이 글자 폭 기반이라 OS·폰트와 무관하게 결정적).

## 4. 대상 월 (정규반 월별 계획서)

`TARGET_MONTH`에 월(숫자)을 넣으면 **정규반 강좌만** 그 달 진도로 잘라
회차·수강기간·수강료를 그 달 기준으로 재계산합니다. 전체 기간 계획서를 만들려면
`None`으로 둡니다. (입력칸 `input()` 대신 변수라 셀 재실행에 안전합니다.)

```python
TARGET_MONTH = 7   # 생성할 월(정규반). 전체 기간 계획서는 None.
```

## 5. 실행

```python
from src.pipeline import run_pipeline

result = run_pipeline(
    input_path="/content/lecture-plan-automation/input",
    output_dir="/content/lecture-plan-automation/output",
    base_year=2026,
    academic_calendar_path="/content/lecture-plan-automation/config/academic_calendar_sample.csv",
    template_path="/content/lecture-plan-automation/templates/강의계획서_마스터템플릿.pptx",
    make_pptx=True,
    target_month=TARGET_MONTH,   # None이면 전체 기간 / 7이면 7월 정규반
)

result
```

> 월별 실행이면 결과 파일명에 대상 연·월이 붙습니다(예: `강의계획서_초안_2026_07_...`).
> 특강 강좌와 그 달에 수업이 없는 정규반은 슬라이드를 만들지 않습니다.

## 6. (선택) 콜랩에서 PPTX 미리보기

생성 자체는 PowerPoint가 필요 없습니다. 콜랩에서 결과를 이미지로 미리 보려면 LibreOffice로 변환합니다.

```python
!apt-get -qq install -y libreoffice >/dev/null
!libreoffice --headless --convert-to pdf --outdir /tmp "{result['pptx_path']}"
!pip -q install pdf2image && apt-get -qq install -y poppler-utils >/dev/null

from pdf2image import convert_from_path
import glob
pdf = glob.glob('/tmp/*.pdf')[0]
for i, img in enumerate(convert_from_path(pdf, dpi=120), 1):
    display(img)   # 강좌별 슬라이드 미리보기
```

## 7. 결과 다운로드

```python
import os
from google.colab import files

for path in [
    result["normalized_xlsx"],
    result["normalized_json"],
    result["validation_report"],
    result["pptx_path"],
]:
    if path and os.path.exists(path):
        files.download(path)
```
