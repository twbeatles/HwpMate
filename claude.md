# HwpMate 프로젝트 지침서 (Claude)

이 문서는 Claude 계열 코딩 에이전트가 HwpMate 리포지토리를 수정할 때 따라야 할 기준을 정리합니다. 현재 유지보수 중심은 `hwptopdf-hwpx_v4.py`가 호출하는 `hwpmate/` 패키지이며, `legacy/hwptopdf-hwpx v3.py`는 레거시 참고용입니다.

## 1. 프로젝트 개요

- 목적: 한글(HWP/HWPX) 문서를 다양한 형식으로 일괄 변환하는 Windows GUI 도구
- 핵심 기술: `PyQt6`, `pywin32`, `PyInstaller`
- 실행 환경: Windows 10/11, Python 3.10+, 한컴오피스 한글 설치
- 배포 엔트리포인트: `hwptopdf-hwpx_v4.py`

## 2. 절대 깨지면 안 되는 로직

### SaveAs 폴백
- `HWPConverter.convert_file`의 2-인자 `SaveAs` 호출 후 3-인자 `SaveAs(..., "")`로 폴백하는 구조를 유지합니다.
- 한글 버전에 따라 COM 인자 수가 달라질 수 있으므로, 단순화하거나 합치지 않습니다.

### COM 초기화
- 메인 스레드와 워커 스레드에서 COM 초기화/해제를 분리합니다.
- `ConversionWorker.run()`의 `pythoncom.CoInitialize()` / `CoUninitialize()` 호출을 제거하지 않습니다.

### 보안 모듈 등록
- `RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")`는 외부 자동화 경고 완화를 위해 유지합니다.
- `SetMessageBoxMode(0x00000001)`도 함께 유지합니다.

### 네이티브 드래그 앤 드롭
- 관리자 권한 환경 호환을 위해 `NativeDropFilter`와 `WM_DROPFILES` 흐름을 유지합니다.
- 폴더 모드에서는 "폴더 1개 드롭 -> 폴더 선택/미리보기 스캔" 흐름을 유지합니다.
- 파일 모드에서만 다중 파일/폴더 스캔 입력을 파일 목록으로 보냅니다.
- Qt 기본 드래그 앤 드롭만으로 되돌리지 않습니다.

### 자동 백업
- 변환 전 `backup/` 폴더에 원본을 복사하는 `_create_backup` 로직은 안전장치입니다.
- 백업 실패는 기록하되 변환 흐름 전체를 무조건 중단시키지 않는 현재 방향을 유지합니다.
- 기본값은 백업 사용(`backup_enabled=True`)이며, 사용자가 끄더라도 변환 흐름은 유지합니다.
- 폴더 재귀 스캔에서는 앱이 만든 하위 `backup/` 폴더를 기본 제외합니다. 사용자가 `backup` 폴더 자체를 직접 선택한 경우만 스캔 대상이 될 수 있습니다.

### 성공 판정과 재시도
- `Open()` 또는 `SaveAs()`가 명시적으로 `False`를 반환하면 실패로 처리합니다.
- 2-인자/3-인자 `SaveAs` 폴백 후 기본 출력 파일이 존재하고 0바이트보다 클 때만 성공으로 집계합니다.
- 실패 자동 재시도는 설정값 `retry_count`를 따르며 기본 1회, 최대 3회입니다.

### 동일 형식 건너뜀과 결과 집계
- `TaskPlanner.build_tasks`는 `PlannedConversion`을 만들고, 동일 형식 입력은 `skipped_tasks`로 분리합니다.
- `ConversionWorker.task_completed`는 `ConversionSummary`를 전달하며, `성공/실패/건너뜀/취소됨` 집계를 분리합니다.
- 동일 형식만 선택된 경우에도 변환 워커를 시작하지 않고 `건너뜀` 전용 결과 다이얼로그를 표시합니다.
- `ResultDialog`와 결과 저장(CSV/JSON/TXT)은 이 집계를 기준으로 동작해야 하며, CSV/JSON에는 `retry_count`, `backup_file`, `backup_error`가 포함됩니다.

### 강제 종료 안전장치
- 강제 종료는 `HWPConverter.kill_owned_processes()`를 통해 앱이 직접 띄운 PID에만 적용합니다.
- 프로세스명 기준 전체 `Hwp.exe` 종료 방식으로 되돌리지 않습니다.

## 3. 코드베이스 구조

- `hwptopdf-hwpx_v4.py`
  - 패키지 진입용 얇은 래퍼
- `hwpmate/`
  - `config_repository.py`, `path_utils.py`, `models.py`: 설정/경로/데이터 모델 (`AppConfig`, `ConversionTask`, `PlannedConversion`, `ConversionSummary`)
  - `services/hwp_converter.py`: HWP COM 래퍼
  - `services/file_selection_store.py`, `services/task_planner.py`: 파일 선택 상태와 작업 계획/건너뜀/출력 충돌 계산
  - `workers/file_scan_worker.py`, `workers/conversion_worker.py`: 비동기 스캔/변환 워커
  - `windows_integration.py`: 관리자 권한 드롭 처리
  - `ui/main_window.py`: 전체 UI 오케스트레이션
  - `ui/theme.py`, `ui/toast.py`, `ui/widgets.py`, `ui/dialogs.py`: UI 컴포넌트 분리 (`PreflightDialog`, `ResultDialog` 포함)
  - `ui/main_window_ui.py`: 메인 윈도우 레이아웃 빌더
- `hwp_converter.spec`
  - PyInstaller 경량 빌드
  - `uac_admin=True` 유지
  - `hiddenimports`와 `EXCLUDES`는 빌드 안정성에 직접 영향
- `pyrightconfig.json`
  - 타입 검사 기준
- `.editorconfig`
  - 인코딩과 줄바꿈 기준

## 4. 수정 원칙

1. 기능 추가보다 기존 자동화 흐름의 호환성을 우선합니다.
2. COM/Qt 경계에서 `None` 가능성과 타입 불일치를 무시하지 않습니다.
3. 문자열과 문서는 `utf-8`을 유지하고, 줄바꿈은 `LF`를 기준으로 맞춥니다.
4. 빌드 관련 변경은 `hwp_converter.spec`와 README를 함께 갱신합니다.
5. 문서에서 지원 형식이나 단축키를 바꾸면 README와 가이드 문서를 같이 수정합니다.

## 5. 검증 기준

최소한 아래 검증은 통과한 뒤 마무리합니다.

```bash
pyright .
```

가능하면 추가로 확인할 것:

- 관리자 권한에서 앱 실행
- 파일/폴더 드롭 동작
- 문서 형식 변환 1건 이상
- 사전 점검 다이얼로그 표시
- 결과 CSV/JSON 저장
- 백업 옵션 및 재시도 횟수 저장/복원
- 동일 형식만 있는 경우의 건너뜀 전용 결과
- 취소 후 종료/강제 종료 흐름
- `pyinstaller hwp_converter.spec` 빌드 여부

## 6. 문서 동기화 체크리스트

- README의 지원 형식, 실행 방법, 빌드 결과 이름이 현재 코드와 일치하는가
- `PROJECT_STRUCTURE_ANALYSIS.md`의 아키텍처 설명과 스냅샷 정보가 현재 구조와 일치하는가
- `update_history.md`에 유지보수 내역이 반영되었는가
- `IMPLEMENTATION_RISK_REVIEW.md`와 `HWP_COM_SMOKE_TEST_CHECKLIST.md`가 현재 구현/검증 기준과 일치하는가
- `.gitignore`가 `backup/`, 빌드 산출물, 캐시를 충분히 제외하는가
- 정적 분석 설정이 실제 코드 상태와 맞는가
