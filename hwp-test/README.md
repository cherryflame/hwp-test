# HWP 5.x 읽기 EXE 검증판

목적은 문서 정리기 본체를 만들기 전에 **Python이 설치되지 않은 Windows PC에서 HWP 5.x 본문 추출이 되는지** 확인하는 것입니다.

## GitHub에서 EXE 만들기
1. 새 GitHub 저장소를 만듭니다.
2. 이 ZIP의 내용물을 저장소 최상위에 그대로 업로드합니다. `.github` 폴더도 반드시 포함합니다.
3. 저장소의 **Actions** 탭 → **Build Windows EXE** → **Run workflow**를 누릅니다.
4. 빌드가 끝난 실행 결과 페이지 아래 **Artifacts**에서 `HWP5_Read_Test_Windows`를 받습니다.
5. 압축을 풀어 `HWP5_Read_Test.exe`를 실행합니다.

## 테스트
`HWP 파일 선택`을 눌러 평소 사용하는 `.hwp`를 엽니다. 본문이 화면에 정상적으로 나오면 1차 검증 성공입니다.

이 프로그램은 파일을 읽기만 하며 저장·수정·삭제하지 않습니다.

## 기술 메모
- HWP/HWPX 읽기: syhwp 0.0.8 (MIT)
- EXE 패키징: PyInstaller
- Windows 빌드: GitHub Actions `windows-latest`
