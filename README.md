# HwpMate

한컴오피스 한글(HWP/HWPX) 문서를 PDF, HWPX, DOCX, ODT, HTML, RTF, TXT와 이미지 형식으로 일괄 변환하는 Windows 전용 GUI 도구입니다. 현재 배포 대상 엔트리포인트는 루트의 얇은 래퍼 `hwptopdf-hwpx_v4.py`이며, 실제 구현은 `hwpmate/` 패키지 아래의 PyQt6 UI와 pywin32 기반 HWP COM 자동화 모듈로 분리되어 있습니다.

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
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
- 변환 시작 전 사전 점검 다이얼로그에서 실행 대상, 건너뜀, 출력 충돌 조정, 입력 파일/출력 폴더 상태, 백업/재시도 설정을 확인할 수 있습니다.
- 한글 COM 허용·보안 창이 뒤에 가려질 수 있어, 변환 연결 중 한글 창 전면화와 안내 토스트를 제공합니다.
- 프로세스 조회는 `tasklist` 대신 Win32 Toolhelp를 사용해 변환 중 콘솔 창이 뜨지 않습니다.
- 한컴 오토메이션 보안승인 모듈(DLL)을 앱이 자동 설치·레지스트리 등록해 파일별 허용 팝업을 억제합니다. 설치 전 SHA-256 무결성을 검증합니다.
- 모듈 준비에 실패할 때만 허용 창이 뜨며, 변환 중 보조로 '모두 허용' 자동 시도를 할 수 있습니다(옵션으로 끌 수 있음). 모듈 등록 성공 시 자동 클릭은 생략됩니다.
- 토스트는 변환 시작·완료·경고 등 짧은 안내를 화면 우측 하단에 띄우며, 배경은 거의 불투명한 슬레이트 톤(`#0f172a`), 본문은 **흰색 굵은 글씨**로 가독성을 맞춥니다.
- PDF·이미지 등 일부 형식은 한컴오피스 인쇄/용지 설정에 따라 결과가 달라질 수 있음을 사전 점검·도움말에 안내합니다.
- 변환 전 원본을 `backup` 폴더에 자동 백업하며, 필요 시 백업을 끌 수 있습니다.
- 실패한 파일은 기본 1회 자동 재시도하며, 재시도 횟수는 0~3회로 조정할 수 있습니다.
- 동일 형식 변환(`HWP->HWP`, `HWPX->HWPX`)은 자동으로 건너뛰고 결과에 별도 집계합니다.
- 다크/라이트 테마, 상태바, 시스템 트레이, **고대비 토스트 알림**(짙은 배경 + 흰색 굵은 글씨, 유형별 강조 테두리)을 포함한 현대적인 UI를 제공합니다.
- 중복 파일 감지, 출력 경로 유효성 검사, 기존 파일 덮어쓰기와 배치 내부 출력 충돌 보호를 지원합니다.
- 이미지/HTML 보조 산출물도 출력 충돌 회피와 성공 판정에 함께 반영합니다.
- 실패 목록 TXT와 전체 결과 CSV/JSON 저장을 지원하며, 결과에는 산출 파일/크기/수정 시각/COM 형식 감사 정보가 포함됩니다.
- 결과 TXT/CSV/JSON과 설정 파일은 임시 파일 작성 후 교체해 부분 저장 위험을 줄입니다.
- 단일 인스턴스 잠금과 변환 중 입력/단축키/드롭 차단으로 중복 실행과 상태 꼬임을 줄입니다.
- 폴더 스캔 대기·작업 계획 중(`is_planning`)에도 시작/입력/드롭 재진입을 막습니다.
- 폴더 미리보기 캐시는 만료·샘플 존재 검증 후 사용하며, 캐시 없으면 비동기 스캔 후 변환합니다(UI 스레드 전체 재스캔 없음).
- 강제 종료는 앱이 직접 띄운 한글 프로세스에만 제한적으로 적용합니다.
- `pyright`·`pytest` 기준 정적 검사와 단위 테스트를 통과하도록 관리합니다.

## 실행 환경

| 항목 | 요구사항 |
|------|----------|
| 운영체제 | Windows 10/11 64-bit |
| Python | 3.10 이상 |
| 한글 | 한컴오피스 한글 2018 이상 |
| 권한 | 관리자 권한 권장 및 사실상 필수 |

## 설치 및 실행

```bash
pip install PyQt6 pywin32
python hwptopdf-hwpx_v4.py
```

- Windows에서 관리자 권한으로 실행해야 HWP COM 자동화와 드래그 앤 드롭이 안정적으로 동작합니다.
- 레거시 tkinter 구현은 `legacy/hwptopdf-hwpx v3.py`에 보관되며, 현재 유지보수와 빌드는 `v4` 기준으로 진행합니다.

## 빌드

```bash
pyinstaller --noconfirm --clean hwp_converter.spec
```

- 실행 파일 이름은 `dist/HWP변환기_v8.7.exe` **단일 파일**입니다. 배포 시 이 exe만 복사하면 됩니다.
- 빌드 전 `hwpmate/resources/security/FilePathCheckerModuleExample.dll` 이 있어야 합니다(없으면 spec 이 실패).
- 한컴 보안승인 모듈 DLL은 exe 안에 포함되며, 최초 변환 시 `%LOCALAPPDATA%\HwpMate\security\` 로 풀린 뒤 레지스트리에 등록됩니다. 사용자 PC에 DLL을 따로 둘 필요는 없습니다.
- `.spec` 은 루트 래퍼 `hwptopdf-hwpx_v4.py` 기준이며 `uac_admin=True` 로 관리자 권한을 요청합니다.

## 개발 품질 기준

```bash
pyright .
pytest
```

- `pyrightconfig.json`을 리포지토리 기준 타입체크 설정으로 사용합니다.
- `.editorconfig`로 `utf-8`, `LF`, 최종 개행 규칙을 고정해 인코딩 및 줄바꿈 혼선을 줄입니다.
- 실제 사용자 데이터와 로그는 리포지토리 바깥 사용자 디렉터리에 저장됩니다.
  - 설정: `~/.hwp_converter_config.json`
  - 로그: `~/.hwp_converter/logs` 또는 `%LOCALAPPDATA%\HwpMate\logs`
  - 단일 인스턴스 잠금: `%LOCALAPPDATA%\HwpMate\HwpMate.lock`
- 실제 한글 COM 스모크는 관리자 권한 PowerShell에서 `python tools/hwp_com_smoke.py --input <샘플.hwp> --format PDF --output-dir <출력폴더>`로 보조 확인할 수 있습니다.

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
├── hwptopdf-hwpx_v4.py          # 배포/실행 래퍼
├── legacy/
│   └── hwptopdf-hwpx v3.py      # 레거시 참고용
├── hwpmate/
│   ├── app.py / bootstrap.py / app_instance.py
│   ├── constants.py / config_repository.py / models.py
│   ├── path_utils.py / logging_config.py
│   ├── windows_integration.py  # 네이티브 DnD, 한글 창 전면화
│   ├── resources/security/     # FilePathChecker 보안 모듈 DLL (빌드 필수)
│   ├── services/
│   │   ├── artifact_policy.py
│   │   ├── hwp_converter.py
│   │   ├── hwp_security_module.py   # DLL 설치·SHA-256·레지스트리
│   │   ├── hwp_security_session.py  # 전면화/자동 클릭 정책
│   │   ├── file_selection_store.py
│   │   └── task_planner.py
│   ├── workers/                # 스캔·변환 QThread
│   └── ui/                     # MainWindow, 컨트롤러, 토스트, 다이얼로그
├── tests/
├── tools/
│   └── hwp_com_smoke.py
├── hwp_converter.spec
├── pyrightconfig.json
├── .editorconfig
├── README.md
├── HWP_COM_SMOKE_TEST_CHECKLIST.md
├── PROJECT_STRUCTURE_ANALYSIS.md
├── PROJECT_AUDIT.md
├── update_history.md
├── claude.md
└── gemini.md
```

## 주의사항

1. 변환 중에는 한글 프로그램을 직접 조작하지 않는 편이 안전합니다.
2. 출력 형식에 따라 한글 설치 버전별 COM 호환 차이가 있을 수 있으므로 `SaveAs` 폴백 로직을 유지해야 합니다.
3. 이미지 변환(`PNG`, `JPG`, `BMP`, `GIF`)과 `HTML`은 한글 설치 버전별 저장 동작이 다를 수 있으며, 앱은 기본 출력 파일과 같은 stem 기반 보조 산출물의 생성/갱신 여부를 성공 기준과 충돌 회피 기준에 함께 사용합니다.
4. 동일 형식 파일은 자동으로 건너뛰며, 결과 창과 결과 리포트에 `건너뜀`으로 표시됩니다.
5. 테스트용 문서를 이 리포지토리 안에서 변환할 경우 `backup/` 폴더가 생성될 수 있으며, 이는 기본적으로 Git 추적 대상이 아닙니다. 폴더 스캔 시 하위 `backup/` 폴더는 기본 제외됩니다.
6. 두 번째 앱 실행은 기존 인스턴스를 보호하기 위해 안내 후 종료됩니다.

## 문서 안내

- [update_history.md](update_history.md): 기능 변화와 유지보수 이력
- [HWP_COM_SMOKE_TEST_CHECKLIST.md](HWP_COM_SMOKE_TEST_CHECKLIST.md): 실제 한글 COM 수동 검증 체크리스트
- [PROJECT_STRUCTURE_ANALYSIS.md](PROJECT_STRUCTURE_ANALYSIS.md): 아키텍처와 확장 포인트 분석
- [PROJECT_AUDIT.md](PROJECT_AUDIT.md): 기능 감사 리포트 (권장안 반영 상태 포함)
- [claude.md](claude.md): Claude 계열 협업 가이드
- [gemini.md](gemini.md): Gemini 계열 협업 가이드

MIT License
