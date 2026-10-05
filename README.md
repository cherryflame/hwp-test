# 문서 중복·유사성 검사기 — EPUB 지원판 v27

기준: 실제 Windows에서 속도와 공백·줄바꿈 표시를 확인한 기본판 v23에서 분기한 EPUB 지원판. v27은 v25의 속도 개선을 유지하면서 EPUB 교차 형식 유사도 산식을 수정했습니다.

지원 형식: HWP 5.x / HWPX / DOCX / TXT / EPUB

## EPUB 지원판 원칙
- EPUB은 META-INF/container.xml → OPF → manifest/spine 순서로 본문 XHTML을 읽습니다.
- EPUB이 없는 HWP/HWPX/DOCX/TXT 비교는 v23의 기존 후보 생성/비교 경로를 유지합니다.
- EPUB 교차 형식은 추출 길이 70% 조건을 강제하지 않습니다.
- EPUB 교차 형식 유사도는 공백·줄바꿈·문단 장식 기호를 점수에서 제외한 문자 흐름의 7문자 shingle Dice로 계산합니다. 장편 전체 문자 SequenceMatcher는 사용하지 않습니다.
- 유사도 계산용 정규화와 상세 표시용 원문을 분리합니다. 상세 비교에서는 공백·줄바꿈 차이를 계속 표시합니다.
- 한 폴더만 등록해도 그 폴더 안의 EPUB↔TXT/HWP/HWPX/DOCX 후보를 찾습니다.
- Windows Explorer 큰 아이콘용 128/256px 리소스를 ICO에 추가했습니다.

## 빌드
GitHub Actions의 `Build Windows EXE`를 실행합니다.
산출물: `문서_중복_유사성_검사기_EPUB_지원판.exe`

원본 문서는 읽기만 하며 삭제·이동·수정하지 않습니다.


## v27 변경
- EPUB 교차형식 상세 비교에 진행형 semantic anchor 재동기화를 적용했습니다.
- 줄바꿈과 문장부호 삭제가 있어도 이후 본문 전체가 변경 블록으로 밀리지 않습니다.
- 전체 문서 문자 SequenceMatcher를 사용하지 않아 v26의 빠른 비교 경로를 유지합니다.
