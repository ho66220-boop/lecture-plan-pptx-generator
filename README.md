# Lecture Plan PPTX Generator

강의계획서 PPTX 제작 과정에서 반복되는 수작업 입력을 줄이기 위한 개인 업무 자동화 도구입니다.

## 목적

- 강사 입력 자료를 표준 엑셀 양식으로 정리
- 엑셀 데이터를 기반으로 강의계획서 PPTX 초안 생성
- 교무팀은 최종 검수와 미세 수정에 집중

## 사용 흐름

1. `input/sample_input.xlsx` 형식에 맞춰 강의 정보를 입력합니다.
2. `template/lecture_plan_template.pptx`를 기준 템플릿으로 사용합니다.
3. `lecture_plan_generator.ipynb`를 Colab에서 실행합니다.
4. 생성된 PPTX 파일을 다운로드한 뒤 최종 검수합니다.

## 폴더 구조

```text
lecture-plan-pptx-generator/
├─ README.md
├─ requirements.txt
├─ lecture_plan_generator.ipynb
├─ input/
├─ template/
├─ output/
└─ notes/
