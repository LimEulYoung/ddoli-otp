# Ddoli OTP

Windows 시스템 트레이 기반 TOTP(Time-based One-Time Password) 관리 프로그램.

## 프로젝트 구조

```
otp_tray.py       # 메인 소스 (단일 파일)
requirements.txt   # Python 의존성
dist/DdoliOTP.exe  # PyInstaller 빌드 결과물
```

## 기술 스택

- Python 3.14, Windows 전용
- pystray: 시스템 트레이 아이콘 (Win32 API 직접 사용)
- tkinter: GUI 다이얼로그
- pyotp: TOTP 코드 생성
- OpenCV (cv2): QR 코드 디코딩
- PIL/Pillow: 아이콘 생성, 화면 캡처 (ImageGrab)
- pyperclip: 클립보드 복사

## 주요 기능

- 트레이 아이콘 좌클릭/우클릭으로 OTP 메뉴 팝업
- OTP 항목 클릭 시 코드 클립보드 복사
- OTP 코드 옆에 남은 시간(초) 표시
- QR 코드 화면 캡처 등록 (드래그 영역 선택 → otpauth:// URI 파싱)
- Secret Key 수동 입력 등록
- OTP 관리 (이름 변경, 삭제)
- Windows 시작 시 자동 실행 (레지스트리 HKCU\...\Run)

## 데이터 저장 경로

`%LOCALAPPDATA%\DdoliOTP\otp_data.json`

## 개발 실행

```bash
venv\Scripts\pythonw.exe otp_tray.py
```

## EXE 빌드

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --name DdoliOTP --icon app_icon.ico otp_tray.py
```

UPX 압축 적용 시 `--upx-dir <upx경로>` 추가.

## 아키텍처 참고

- DPI awareness: `SetProcessDpiAwareness(2)` (Per-Monitor V2). QR 캡처 좌표 정확도에 필수.
- pystray 내부 API: `_message_handlers`, `_update_menu()`, `_menu_handle` 등 private API 사용.
  메뉴 교체 후 반드시 `icon._update_menu()` 호출해야 Win32 HMENU가 갱신됨.
- QR 캡처 시 오버레이를 `withdraw()` 후 `ImageGrab.grab()` 해야 오버레이가 캡처에 포함되지 않음.
- 모든 GUI 다이얼로그는 `threading.Thread(daemon=True)`로 실행 (pystray 메시지 루프 블로킹 방지).
