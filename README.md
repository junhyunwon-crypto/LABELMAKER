# B-FV4T ZMODE Windows 자동 빌드

이 버전은 Windows GDI와 Toshiba TPCL 드라이버의 이미지 렌더링을 사용하지
않습니다. 완성된 320 x 400 라벨을 ZPL 그래픽으로 변환한 뒤 Windows 프린터
대기열에 RAW로 직접 전송합니다.

## 프린터 설정

프린터 본체를 `ZMODE`로 설정하고 전원을 껐다가 다시 켭니다.

## GitHub Actions 빌드

1. 이 ZIP의 내용물을 GitHub 저장소 최상위에 업로드합니다.
2. `.github/workflows/build-windows.yml`이 저장소에 있는지 확인합니다.
3. **Actions → Build Windows EXE → Run workflow**를 실행합니다.
4. 완료 후 **Artifacts → BarcodeMaker_BFV4T_ZMODE**를 받습니다.
5. 압축 안의 `BarcodeMaker_BFV4T_ZMODE.exe`를 프린터 PC에서 실행합니다.

프린터 PC에는 Python이 필요하지 않습니다. 앱에서 선택하는 Windows 프린터
대기열은 실제 B-FV4T USB 포트에 연결되어 있어야 합니다.
