# HwpMate 프로젝트 지침서 (Gemini)

이 문서는 **HwpMate (v8.6+)** 프로젝트를 유지보수하거나 기능을 확장할 때 따라야 할 핵심 지침을 담고 있습니다.

## 1. 프로젝트 개요
**Windows용 한글(HWP) 파일 일괄 변환 도구**입니다.
Python의 `pywin32`를 사용하여 한글 오피스(Hwp Automation)를 제어하며, `PyQt6`로 현대적인 GUI를 제공합니다.

### 핵심 목표
- **안정성**: COM 객체 연결 실패 최소화 및 예외 처리
- **사용성**: 드래그 앤 드롭, 직관적인 UI, 관리자 권한 호환성
- **확장성**: PDF 외 HWPX, DOCX 등 다양한 포맷 지원

---

## 2. ⚠️ 절대 변경 금지 영역 (Mission Critical)

다음 로직은 수많은 시행착오 끝에 정착된 안정화 코드입니다. **명확한 이유 없이 수정하거나 리팩토링하지 마십시오.**

### 2.1. 변환 엔진 (`HWPConverter.convert_file`)
한글 버전에 따라 `SaveAs` 메서드의 인자가 다르기 때문에 **이중 try-except 구조**를 필수적으로 유지해야 합니다.

```python
# 기존 코드 구조 유지 필수
try:
    # 시도 1: 구버전 호환 (2개 인자)
    self.hwp.SaveAs(output_str, save_format)
except Exception as e1:
    try:
        # 시도 2: 신버전 호환 (3개 인자: 경로, 포맷, 빈 문자열)
        # 빈 문자열("") 인자가 없으면 일부 버전에서 TypeError 발생
        self.hwp.SaveAs(output_str, save_format, "")
    except Exception as e2:
        # 최종 실패 처리
        ...
```

### 2.2. 보안 모듈 등록
최신 한글 버전에서는 자동화 스크립트 실행 시 보안 확인 팝업이 뜹니다. 이를 방지하기 위한 모듈 등록 코드를 제거하지 마십시오.
```python
self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
```

### 2.3. 스레딩 및 COM 초기화
`ConversionWorker` 스레드 내에서 `pythoncom.CoInitialize()` 호출은 필수입니다. 
- 메인 스레드와 별도로 워커 스레드에서 COM 객체를 생성하고 해제해야 차단(Freezing) 현상이 없습니다.

### 2.4. 자동 백업 (`_create_backup`)
사용자 데이터를 보호하기 위해 변환 전 `backup` 폴더에 원본 사본을 생성합니다. **이 로직은 선택 사항이 아닌 필수 안전장치**로 유지되어야 합니다.

---

## 3. 기술 스택 및 아키텍처 (Deep Dive)

| 구분 | 기술 | 비고 |
|------|------|------|
| **언어** | Python 3.9+ | |
| **GUI** | PyQt6 | Qt Designer 미사용, 100% 코드로 작성 |
| **자동화** | pywin32 (win32com) | HWP Automation API 사용 |
| **빌드** | PyInstaller | `hwp_converter.spec` 사용 |

### 3.1. 파일 구조 상세
- **`hwptopdf-hwpx_v4.py`**: 
  - `MainWindow`: 메인 GUI, 스크롤 가능 영역(`QScrollArea`) 기반.
  - `NativeDropFilter`: 관리자 권한 드롭 지원을 위한 `ctypes` 후킹.
  - `ConversionWorker`: 실제 변환을 담당하는 `QThread`.
- **`hwp_converter.spec`**: 
  - **UAC 설정**: `uac_admin=True`로 빌드된 exe가 항상 관리자 권한을 요구하도록 설정됨.
  - **최적화**: `numpy`, `pandas`, `QtWebEngine` 등 불필요한 라이브러리를 명시적으로 제외(`EXCLUDES`)하여 용량 최적화.
  - **Hidden Imports**: `win32com.client`, `pythoncom` 등이 명시되어야 실행 오류 방지.

### 3.2. 설정 관리 (`~/.hwp_converter_config.json`)
앱은 종료 시 다음 정보를 저장하고 시작 시 복원합니다:
- `theme` (dark/light), `mode` (folder/files)
- `format` (PDF/HWPX/DOCX)
- `include_sub` (하위 폴더 포함 여부)
- `same_location` (저장 위치)

---

## 4. UI/UX 구현 가이드

이 프로젝트는 "현대적이고 프리미엄한 느낌"을 지향합니다.

### 4.1. 스타일링 규칙
- **Pure CSS**: Qt Style Sheet(QSS)를 사용하여 `ThemeManager` 클래스 내에 정의합니다.
- **색상 팔레트 (다크 테마 기준)**:
  - 배경: `#1a1a2e` (Deep Blue)
  - 패널: `#16213e`
  - 포인트: `#e94560` (Pinkish Red) - 버튼, 강조
  - 보더: `#0f3460`

### 4.2. 특수 컴포넌트
- **FormatTabs**: `QTabWidget`을 사용하여 '문서 변환'과 '이미지 변환'으로 카테고리를 나누었습니다.
- **FormatCard**: 단순 라디오 버튼이 아닌, 아이콘과 설명을 포함한 클릭 가능한 카드 UI입니다. `selected` 속성에 따라 스타일이 변경됩니다.
- **ToastWidget**: 우측 하단에서 올라오는 비침습적 알림입니다. `ToastManager`가 화면에 표시되는 알림의 위치를 스택처럼 관리합니다.
- **DropArea**:
  - 관리자 권한 실행 시, Qt의 기본 드래그 앤 드롭(`dragEnterEvent`)이 UIPI(User Interface Privilege Isolation)에 의해 차단될 수 있습니다.
  - 이를 우회하기 위해 `NativeDropFilter`가 `WM_DROPFILES` 메시지를 직접 가로채 처리합니다.
  - `EnumChildWindows`를 사용하여 메인 윈도우뿐만 아니라 모든 자식 위젯에도 드롭 필터를 등록합니다.

---

## 5. 개발 및 디버깅 팁

1. **관리자 권한 필수**: 
   - 한글 오피스 자동화 인터페이스는 관리자 권한 프로세스에서 실행되어야 정상적으로 파일 접근이 가능합니다.
   - 개발 시 IDE도 **관리자 권한**으로 실행하십시오.
2. **파일 덮어쓰기 로직**:
   - `overwrite` 옵션이 꺼져있으면, 파일명 뒤에 `(1)`, `(2)` 등을 자동으로 붙여 원본 손실을 방지합니다. (`_adjust_output_paths` 메서드 참고)
3. **배치 처리 성능**:
   - 파일 목록 추가 시 `file_table.blockSignals(True)`를 사용하여 대량의 파일 추가 시 UI 멈춤을 방지합니다.

## 6. 버전 관리 및 배포
- 기능 추가 시 `update_history.md`에 기록하고 코드 상단 `VERSION` 상수를 변경하십시오.
- 배포 빌드 시 반드시 `pyinstaller hwp_converter.spec` 명령어를 사용하여 `spec` 파일의 설정을 반영해야 합니다.
