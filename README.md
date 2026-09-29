# 문서 중복·유사성 검사기 — EPUB 지원판

Windows용 문서 중복·유사성 검사기입니다. 확정 기본판 v23을 기준으로 EPUB 지원만 확장한 버전입니다.

## 지원 형식
- HWP 5.x
- HWPX
- DOCX
- TXT
- EPUB

EPUB은 표준 `META-INF/container.xml` → OPF → spine 읽기 순서를 따라 XHTML/HTML 본문을 추출합니다. nav 전용 문서는 본문 중복을 줄이기 위해 제외합니다. DRM/비표준 구조 또는 파싱할 수 없는 EPUB은 읽기 실패로 표시될 수 있습니다.

## 유지된 기본판 기준
- 폴더 후보 축소 방식 유지(O(n²) 전수 비교로 회귀하지 않음)
- 폴더 결과 필터/CSV/결과→A/B 비교 유지
- A/B 비교 진행률 및 sequence-id 유지
- 실제 문구 변경은 빠른 줄/블록 단위 표시
- 공백·줄바꿈만 다른 경우 파란색으로 `·` / `↵` 위치 표시
- 원본 파일 삭제·이동·수정 기능 없음

## 빌드
GitHub Actions의 `Build Windows EXE`를 실행합니다. 산출물 이름은 `문서_중복_유사성_검사기_EPUB_지원판_Windows`, EXE 이름은 `문서_중복_유사성_검사기_EPUB_지원판.exe`입니다.
