# 빌드/패키징 변경 사항

기능 비교 로직은 변경하지 않았습니다.

- PyInstaller `--onefile` → `--onedir`
- `--noupx` 명시
- Windows PE 버전 정보 추가 (`version_info.txt`)
- 16/24/32/48/64/128/256 멀티사이즈 ICO 재생성
- Tk 아이콘 우선순위를 64→48→32→24→16으로 변경
- Windows AppUserModelID 지정으로 작업표시줄의 Python/Tk 기본 아이콘 사용 가능성 완화
- GitHub Actions에서 ZIP + SHA-256 생성

주의: 이 변경은 백신 오탐 가능성을 낮추기 위한 것이며, 특정 백신에서 탐지가 절대 발생하지 않는다고 보장하지 않습니다. 새 빌드가 Defender에서 다시 탐지되면 해당 EXE/ZIP을 배포하지 말고 탐지 결과를 재확인해야 합니다.
