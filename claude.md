# HwpMate 프로젝트 지침서 (Claude)

이 문서는 **HwpMate (v8.6+)** 프로젝트의 구조, 핵심 로직, 그리고 개발 규칙을 설명합니다. AI 어시스턴트로서 본 프로젝트를 수정할 때 이 가이드를 엄격히 준수하십시오.

## 1. Context & Architecture

### 1.1. 프로젝트 개요
- **목적**: 한글(HWP/HWPX) 파일을 PDF, DOCX 등으로 일괄 변환
- **환경**: Windows 10/11, Python 3.9+, 한글 2018 이상
- **프레임워크**: PyQt6 (GUI), pywin32 (COM Automation)

### 1.2. 주요 파일 및 역할
- `hwptopdf-hwpx_v4.py`: Entry point. `MainWindow`, `ConversionWorker`, `NativeDropFilter` 등 모든 핵심 클래스가 포함됨.
- `hwp_converter.spec`: PyInstaller 빌드 설정 파일. 용량 최적화를 위한 Exclude 목록과 관리자 권한(`uac_admin=True`) 설정이 포함됨.
- `ThemeManager`: CSS 기반 테마(Dark/Light) 관리.

---

## 2. 🛡️ Critical Logic (변경 주의)

아래 명시된 로직은 프로젝트의 안정성을 보장하는 핵심 코드입니다. **기능 추가 시에도 기존 동작을 파괴하지 않도록 주의하십시오.**

- **이중 SaveAs 전략 (`HWPConverter`)**:
  - `SaveAs(path, format)` 실패 시 `SaveAs(path, format, "")`를 호출하는 폴백 메커니즘 필수.
  - 한글 버전에 따른 인자 개수 차이(`TypeError`)를 해결하기 위함입니다.

- **보안 모듈 및 팝업 제어**:
  - `RegisterModule("FilePathCheckDLL", ...)`: 보안 경고창 억제.
  - `SetMessageBoxMode(0x00000001)`: 확인 팝업 자동 넘김.

- **스레드 안정성 (`ConversionWorker`)**:
  - `QThread` run 메서드 진입 시 `pythoncom.CoInitialize()` 필수 호출.
  - 이를 누락하면 COM 객체 호출이 실패하거나 애플리케이션이 멈춥니다.

- **자동 백업 시스템 (`_create_backup`)**:
  - 파일 변환 전 원본의 백업본을 `backup` 폴더에 생성합니다.
  - 이 과정에서 실패하더라도(예: 권한 문제) 변환 작업 자체는 중단되지 않도록 `try-except`로 감싸져 있습니다.

---

## 3. UI/UX & Implementation Details

### 3.1. Design Philosophy
- **Modern & Premium**: 윈도우 기본 UI 대신 커스텀 스타일(QSS)을 최대한 활용.
- **Tabbed Interface**: 포맷이 많아짐에 따라 `QTabWidget`을 사용하여 '문서 변환'과 '이미지 변환'으로 카테고리를 분리.
- **Dark Mode First**: 기본값은 다크 모드(`#1a1a2e`).

### 3.2. Advanced Features
- **Native Drag & Drop**:
  - 관리자 권한으로 실행되는 앱은 일반 권한의 탐색기에서 드래그 앤 드롭을 받을 수 없습니다 (Windows 보안 정책).
  - 이를 해결하기 위해 `NativeDropFilter` 클래스가 `WM_DROPFILES` 메시지를 후킹합니다.
  - `EnumChildWindows`를 통해 메인 윈도우의 모든 자식 위젯까지 등록하여 드롭 영역을 확장합니다.

- **UI Event Optimization**:
  - 파일 목록에 수백 개의 파일을 추가할 때 성능 저하를 막기 위해 `blockSignals(True)` 패턴을 사용합니다.
  - `QScrollArea`를 루트 위젯으로 사용하여 낮은 해상도에서도 UI가 잘리지 않도록 설계되었습니다.

- **Configuration Persistency**:
  - 사용자 홈 디렉토리의 `.hwp_converter_config.json` 파일에 마지막 사용 설정을 저장합니다.
  - 저장 항목: 테마, 폴더 경로, 변환 옵션 등.

---

## 4. Development Rules

1. **단일 파일 유지**: 배포 용이성을 위해 가능한 `hwptopdf-hwpx_v4.py` 하나의 파일에 코드를 작성합니다.
2. **관리자 권한**: 이 프로그램은 **반드시 관리자 권한**으로 실행되어야 합니다. 개발 및 테스트 시에도 이를 준수하십시오.
3. **PyInstaller 호환성**: 
   - `spec` 파일에 정의된 `hiddenimports`(`win32com.client`, `pythoncom` 등) 의존성을 유의하십시오.
   - 외부 리소스(이미지 등)는 코드 내장 방식을 선호합니다.
