# 문서 중복·유사성 검사기 — EPUB 지원판

기준: 실제 Windows에서 속도와 공백·줄바꿈 표시를 확인한 기본판 v23에서 분기.

지원 형식: HWP 5.x / HWPX / DOCX / TXT / EPUB

## EPUB 지원판 원칙
- EPUB은 META-INF/container.xml → OPF → manifest/spine 순서로 본문 XHTML을 읽습니다.
- EPUB이 없는 HWP/HWPX/DOCX/TXT 비교는 v23의 기존 후보 생성/비교 경로를 유지합니다.
- EPUB 교차 형식은 추출 길이 70% 조건을 강제하지 않습니다.
- EPUB 교차 형식 유사도는 장편 전체 문자 SequenceMatcher 대신 단어 5-gram 기반의 빠른 비교를 사용합니다.
- 상세 비교는 단일 개행을 문장 경계로 삼지 않고 문장/빈 줄 단위로 정렬하여, 같은 문장의 공백·줄바꿈 차이는 파란색으로 표시하고 실제 문구 변경만 노란색으로 표시합니다.
- 한 폴더만 등록해도 그 폴더 안의 EPUB↔TXT/HWP/HWPX/DOCX 후보를 찾도록 설계되어 있습니다.

## 빌드
GitHub Actions의 `Build Windows EXE`를 실행합니다.
산출물: `문서_중복_유사성_검사기_EPUB_지원판.exe`

원본 문서는 읽기만 하며 삭제·이동·수정하지 않습니다.
