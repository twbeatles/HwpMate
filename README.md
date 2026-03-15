# HwpMate

한컴오피스 한글(HWP/HWPX) 문서를 PDF, HWPX, DOCX, ODT, HTML, RTF, TXT와 이미지 형식으로 일괄 변환하는 Windows 전용 GUI 도구입니다. 현재 배포 대상 엔트리포인트는 `hwptopdf-hwpx_v4.py`이며, PyQt6 UI와 pywin32 기반 HWP COM 자동화를 사용합니다.

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![PyQt6](https://img.shields.io/badge/PyQt6-6.x-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows_10/11-lightgrey.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

## 주요 기능

| 분류 | 지원 형식 |
|------|-----------|
| 문서 | `PDF`, `HWP`, `HWPX`, `DOCX`, `ODT`, `HTML`, `RTF`, `TXT` |
| 이미지 | `PNG`, `JPG`, `BMP`, `GIF` |

- 폴더 일괄 변환과 파일 개별 선택을 모두 지원합니다.
- 관리자 권한 환경에서도 동작하는 네이티브 드래그 앤 드롭을 제공합니다.
- 변환 전 원본을 `backup` 폴더에 자동 백업합니다.
- 다크/라이트 테마, 상태바, 시스템 트레이, 토스트 알림을 포함한 현대적인 UI를 제공합니다.
- 중복 파일 감지, 출력 경로 유효성 검사, 덮어쓰기 방지 번호 부여를 지원합니다.
- `pyright` 기준 정적 타입 검사를 통과하도록 관리합니다.

## 실행 환경

| 항목 | 요구사항 |
|------|----------|
| 운영체제 | Windows 10/11 64-bit |
| Python | 3.9 이상 |
| 한글 | 한컴오피스 한글 2018 이상 |
| 권한 | 관리자 권한 권장 및 사실상 필수 |

## 설치 및 실행

```bash
pip install PyQt6 pywin32
python hwptopdf-hwpx_v4.py
```

- Windows에서 관리자 권한으로 실행해야 HWP COM 자동화와 드래그 앤 드롭이 안정적으로 동작합니다.
- 레거시 tkinter 구현은 `hwptopdf-hwpx v3.py`에 남아 있지만, 현재 유지보수와 빌드는 `v4` 기준으로 진행합니다.

## 빌드

```bash
pyinstaller hwp_converter.spec
```

- 실행 파일 이름은 `HWP변환기_v8.6.exe`입니다.
- `.spec` 파일은 `hwptopdf-hwpx_v4.py`를 기준으로 경량 빌드되도록 설정되어 있습니다.
- `uac_admin=True`가 설정되어 있어 배포 실행 파일은 관리자 권한 승격을 요청합니다.

## 개발 품질 기준

```bash
pyright .
```

- `pyrightconfig.json`을 리포지토리 기준 타입체크 설정으로 사용합니다.
- `.editorconfig`로 `utf-8`, `LF`, 최종 개행 규칙을 고정해 인코딩 및 줄바꿈 혼선을 줄입니다.
- 실제 사용자 데이터와 로그는 리포지토리 바깥 사용자 홈 디렉터리 아래에 저장됩니다.

## 단축키

| 단축키 | 동작 |
|--------|------|
| `Ctrl+O` | 파일 추가 |
| `Ctrl+Shift+O` | 폴더 선택 |
| `Ctrl+Enter` | 변환 시작 |
| `Esc` | 변환 취소 |
| `Delete` | 선택 파일 제거 |
| `Ctrl+Delete` | 전체 파일 제거 |
| `F1` | 프로그램 정보 |

## 프로젝트 구조

```text
HwpMate/
├── hwptopdf-hwpx_v4.py
├── hwptopdf-hwpx v3.py
├── hwp_converter.spec
├── pyrightconfig.json
├── .editorconfig
├── README.md
├── update_history.md
├── claude.md
└── gemini.md
```

## 주의사항

1. 변환 중에는 한글 프로그램을 직접 조작하지 않는 편이 안전합니다.
2. 출력 형식에 따라 한글 설치 버전별 COM 호환 차이가 있을 수 있으므로 `SaveAs` 폴백 로직을 유지해야 합니다.
3. 테스트용 문서를 이 리포지토리 안에서 변환할 경우 `backup/` 폴더가 생성될 수 있으며, 이는 기본적으로 Git 추적 대상이 아닙니다.

## 문서 안내

- [update_history.md](update_history.md): 기능 변화와 유지보수 이력
- [claude.md](claude.md): Claude 계열 협업 가이드
- [gemini.md](gemini.md): Gemini 계열 협업 가이드

MIT License
