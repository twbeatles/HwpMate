# HwpMate 프로젝트 구조 정밀 분석 (기능 확장용)

## 1. 분석 범위
- 분석 일자: 2026-02-27
- 보강 일자: 2026-03-10
- 추가 보강 일자: 2026-03-18
- v8.7 반영 일자: 2026-04-27
- 기능 리스크 보강 일자: 2026-05-12
- MainWindow 컨트롤러 리팩토링 반영 일자: 2026-06-10
- 대상 저장소: `D:\twbeatles-repos\HwpMate`
- 분석 목적: "다양한 기능 추가"를 위한 현재 구조, 제약, 확장 포인트 파악

## 2. 참고한 문서
- `README.md`
- `claude.md`
- `gemini.md`
- `update_history.md`
- `HWP_COM_SMOKE_TEST_CHECKLIST.md`
- `hwp_converter.spec`
- `hwptopdf-hwpx_v4.py` (루트 엔트리포인트 래퍼)
- `hwpmate/` (현행 메인 코드)
- `legacy/hwptopdf-hwpx v3.py` (레거시 참고)

## 3. 저장소 구조 스냅샷

| 파일 | 상태 | 역할 |
|---|---|---|
| `hwptopdf-hwpx_v4.py` | 유지 | 패키지 진입용 얇은 래퍼 |
| `hwpmate/` | 모듈 분리 | 현재 메인 애플리케이션 (GUI + 변환엔진 + 워커 + DnD + 설정) |
| `legacy/hwptopdf-hwpx v3.py` | 참고용 | 레거시 tkinter 기반 버전 |
| `hwp_converter.spec` | 추적 | PyInstaller 빌드 설정 (경량화, `uac_admin=True`) |
| `tools/hwp_com_smoke.py` | 스크립트 | 실제 HWP COM 변환 보조 스모크 |
| `pyrightconfig.json` | 추적 | Pylance/pyright 공용 정적 분석 설정 |
| `.editorconfig` | 추적 | UTF-8/LF 편집 규칙 |
| `README.md` | 추적 | 사용자 관점 기능/설치/사용법 |
| `HWP_COM_SMOKE_TEST_CHECKLIST.md` | 추적 | 실제 한글 COM 수동 검증 체크리스트 |
| `claude.md` | 추적 | 핵심 로직 보존 지침(변경 주의사항) |
| `gemini.md` | 추적 | 유지보수/확장 지침(절대 변경 금지 영역 포함) |
| `update_history.md` | 추적 | 버전 이력, 기술적 문제 해결 기록 |

현재 구조는 `hwpmate/` 패키지 기준의 모듈 분리 아키텍처이며, 루트 래퍼와 기존 배포 흐름은 유지됩니다.

## 4. 아키텍처 개요

### 4.1 실행 진입 흐름
1. `main()` 실행
2. `pywin32` 존재 확인
3. 관리자 권한(`is_admin`) 확인
4. `enable_drag_drop_for_admin()` 호출
5. `QApplication` 생성 후 `MainWindow` 실행
6. `showEvent()`에서 `NativeDropFilter` 설치 및 자식 윈도우까지 WM_DROPFILES 등록

### 4.2 주요 클래스/모듈 역할

| 구성요소 | 핵심 역할 |
|---|---|
| `config_repository.py` | 설정 저장/로드와 JSON 손상 백업 처리 |
| `path_utils.py` | 경로 정규화, 권한 검사, 지원 파일 스캔 |
| `models.py` | `AppConfig`, `ConversionTask`, `PlannedConversion`, `ConversionSummary`, `FormatSpec` 데이터 모델 |
| `services/file_selection_store.py` | 순서 유지 + 대소문자 비민감 중복 제거 |
| `services/task_planner.py` | 모드별 작업 생성, 동일 형식 건너뜀 분리, 출력 충돌 해소 |
| `services/hwp_converter.py` | COM 연결/문서 열기/SaveAs/정리 담당 |
| `workers/file_scan_worker.py` | 파일/폴더를 비동기 배치 스캔 |
| `workers/conversion_worker.py` | 작업 리스트 순차 변환, 백업, 취소/요약 집계, 안전한 강제 종료 위임 |
| `windows_integration.py` | 관리자 권한에서도 동작하는 네이티브 DnD 처리 |
| `ui/theme.py`, `ui/toast.py`, `ui/widgets.py`, `ui/dialogs.py` | 테마/토스트/위젯/사전 점검/결과 다이얼로그 |
| `ui/main_window.py` | `MainWindow` import 경로를 유지하는 조립 루트와 호환 래퍼 |
| `ui/main_window_ui.py` | 콜백 객체 기반 메인 윈도우 레이아웃 빌더 |
| `ui/main_window_controllers/state.py` | `MainWindowState` 런타임 상태 모델 |
| `ui/main_window_controllers/appearance.py` | 테마, 포맷 선택, 모드/출력 UI 활성 상태 |
| `ui/main_window_controllers/file_selection.py` | 파일/폴더 선택, 파일 테이블, 비동기 스캔 수명주기 |
| `ui/main_window_controllers/conversion.py` | 작업 계획, 사전 점검, 변환 워커, 결과/취소 처리 |
| `ui/main_window_controllers/native_drop.py` | 네이티브 WM_DROPFILES 초기화와 모드별 드롭 처리 |
| `ui/main_window_controllers/lifecycle.py` | 메뉴, 단축키, 트레이, 설정 저장, 종료 이벤트 |

### 4.3 데이터/상태 모델
- 설정 파일: `%USERPROFILE%\.hwp_converter_config.json`
- 기본 설정 키:
  - `config_version`, `theme`, `mode`, `format`, `include_sub`, `same_location`, `overwrite`
  - `backup_enabled`, `retry_count`
- 추가 저장 키:
  - `folder_path`, `output_path`, `last_folder`, `last_output`
- 런타임 주요 상태:
  - `self.file_list`, `self._file_set`
  - `MainWindowState.plan`, `tasks`, `last_summary`, `worker`, `scan_worker`
  - `MainWindowState.is_converting`, `scan_mode`, `force_kill_pending`, `close_after_worker`, `selected_format`
  - `MainWindow`의 기존 underscore 속성은 테스트/외부 호환용 property 래퍼로 유지

## 5. 실제 동작 플로우

### 5.1 파일 수집 플로우
1. 사용자 입력 (파일 선택, 폴더 선택, 네이티브 드롭)
2. `FileSelectionController.start_scan()`으로 `FileScanWorker` 실행
3. 배치별 `FileSelectionController.on_scan_batch_found()` 호출
4. `append_files_batch()`에서 중복 제거 + 테이블 렌더링
5. `on_scan_finished()`에서 상태 라벨 갱신

### 5.2 변환 플로우
1. `ConversionController.start_conversion()`
2. `collect_tasks()`로 `PlannedConversion` 생성
3. 필요 시 `adjust_output_paths()`로 충돌 회피 수 계산
4. `PreflightDialog`로 실행 대상/건너뜀/경고/입력·출력 차단 오류 확인
5. 실행 대상 없이 동일 형식 건너뜀만 있으면 즉시 `ResultDialog` 표시
6. 실행 대상이 있으면 `ConversionWorker` 시작
7. `ConversionWorker.run()` 내부
   - 워커 스레드 `pythoncom.CoInitialize()`
   - `HWPConverter.initialize()` 및 보안 모듈/PID 추적 경고 수집
   - 파일별 선택적 `_create_backup()` 후 `convert_file()`
   - 실패 시 설정된 횟수만큼 재시도
   - 취소 시 남은 task를 `취소됨`으로 마킹
   - 산출 파일/크기/수정 시각/COM 형식을 기록한 `ConversionSummary` 생성
8. `task_completed` 시그널 수신 후 `ResultDialog` 표시
9. 필요 시 실패 TXT / 결과 CSV·JSON 저장
10. `on_worker_finished()`에서 UI/시그널/상태 정리 및 종료 대기 처리

## 6. 기능 추가 시 반드시 지켜야 할 핵심 제약
(출처: `claude.md`, `gemini.md`, 코드 본문)

1. SaveAs 이중 전략 유지
- 2-파라미터 실패 시 3-파라미터(`""`) 재시도 로직 필수.
- `Open()`/`SaveAs()`가 명시적으로 `False`를 반환하면 실패로 처리.
- 출력 산출물이 새로 생성/갱신됐고 0바이트 초과인지 검증 유지.
- 이미지/HTML 계열은 같은 stem 기반 보조 산출물도 함께 수집.

2. 보안/팝업 제어 유지
- `RegisterModule("FilePathCheckDLL", ...)`
- `SetMessageBoxMode(0x00000001)`

3. 스레드 COM 초기화 유지
- 워커 스레드 진입 시 `pythoncom.CoInitialize()` 필수.

4. 자동 백업 로직 유지
- 백업 실패가 전체 변환 실패로 이어지지 않게 현재 방식을 유지.
- 백업명 충돌 방지를 위한 마이크로초/일련번호 전략 유지.
- 사용자가 백업을 끄는 경우를 제외하고 기본값은 백업 사용.
- 폴더 재귀 스캔에서 하위 `backup/` 폴더를 제외.

5. 관리자 권한 + 네이티브 DnD 경로 유지
- UIPI 우회용 WM_DROPFILES 처리 삭제 금지.
- 긴 경로 대응용 동적 버퍼 할당 유지.

6. 동일 형식 건너뜀/결과 요약 유지
- 동일 형식은 오류가 아니라 `건너뜀`으로 집계.
- 결과 화면과 결과 저장 파일은 `성공/실패/건너뜀/취소됨` 집계를 일관되게 사용.
- 동일 형식만 선택된 경우도 결과 저장 가능해야 함.
- 결과 CSV/JSON에는 `created_files`, `output_size`, `output_mtime`, `save_format`, `progid_used` 감사 필드를 유지.

7. 강제 종료 범위 제한 유지
- 시스템 전체 한글 프로세스 종료로 되돌리지 말고, 앱이 추적한 PID만 종료.

8. 배포 제약 고려
- `hwp_converter.spec`의 `hiddenimports`, `uac_admin=True`, excludes 전략을 손상시키지 말 것.

## 7. 확장성 평가

### 7.1 강점
- 비동기 스캔/변환 스레드 분리 완료.
- 변환 포맷 메타데이터(`FORMAT_TYPES`) 중심 확장 구조.
- `MainWindow`는 조립 루트로 축소되고 UI 런타임 책임이 컨트롤러별로 분리됨.
- 레이아웃 빌더는 `MainWindowCallbacks`로 시그널 연결을 주입받아 직접 underscore 메서드에 덜 결합됨.
- 변환 안정성 관련 폴백/예외 처리 경험치가 코드에 축적됨.

### 7.2 한계
- 컨트롤러가 Qt 위젯을 직접 다루므로 GUI 경계의 수동/오프스크린 검증은 계속 필요합니다.
- 실제 HWP COM 경로는 설치 환경 의존성이 커서 자동 테스트와 별도로 수동 검증 비중이 높습니다.
- 설정 스키마 마이그레이션은 기본값 병합 후 타입/범위 정규화를 수행하며, v8.7에서 `config_version=2`입니다.
- 향후 기능 추가 시 새 컨트롤러 책임 경계를 넘는 상호 호출이 늘어나지 않도록 주의해야 합니다.

## 8. 추천 기능 추가 항목 (우선순위)

### 8.1 단기(낮은 리스크)
1. 변환 프리셋 저장/불러오기
- 예: "PDF-사내공유", "DOCX-검수용"
- 영향 영역: 설정 로드/저장, UI(프리셋 콤보), `_save_settings`, `_start_conversion`

2. 출력 파일명 템플릿
- 예: `{name}_{date}` `{name}_{format}`
- 영향 영역: `_collect_tasks`, `_adjust_output_paths`

3. 백업 보존 정책 확장
- 예: 오래된 백업 자동 정리, 지정 폴더 백업
- 영향 영역: 설정, UI 옵션, `_create_backup`

### 8.2 중기(중간 리스크)
5. 작업 큐 저장/복원
- 대량 작업 중단 후 재실행용 큐 파일
- 영향 영역: `ConversionTask` 직렬화, `MainWindow`, 설정/파일 I/O

6. 변환 전 검증 대시보드
- "권한 없음", "경로 문제", "지원 안 되는 확장자" 사전 표시
- 영향 영역: `_collect_tasks`, 유틸 함수, UI 다이얼로그

7. 출력 후 후처리 훅
- 예: 변환 후 폴더 자동 열기, 파일 압축(zip), 로그 첨부
- 영향 영역: `_on_task_completed`, 새로운 후처리 서비스 함수

8. 사용자 로그 뷰어 탭
- Rotating 로그를 GUI에서 필터 조회
- 영향 영역: UI 탭 추가, 로그 파일 read-only 파서

### 8.3 장기(높은 리스크)
9. 플러그인형 포맷/후처리 확장
- 현재는 `FORMAT_TYPES` 고정 구조. 동적 확장으로 전환 시 설계 변경 필요.

10. 멀티 프로세스 변환 분산
- HWP COM 특성상 프로세스 안정성/라이선스/충돌 검증 선행 필요.

## 9. 안전한 구조 개선 제안
현재 패키지 분리 이후에도 아래 개선 여지는 남아 있습니다:

1. 컨트롤러 간 의존 방향 유지
- `MainWindow`를 경유한 위임은 호환용으로 유지하되, 신규 기능은 해당 책임 컨트롤러 안에서 먼저 닫히도록 설계

2. UI 액션 서비스화
- 폴더 열기, 실패 목록 내보내기, 사용자 알림 정책을 더 작은 서비스로 분리 가능

3. GUI/COM 검증 보강
- Qt 상호작용 테스트와 `tools/hwp_com_smoke.py`는 추가됐으므로, 실제 HWP COM 설치 환경에서는 스크립트와 수동 점검을 함께 사용할 수 있음

4. 빌드 자동화
- `pyright`, `pytest`, `pyinstaller`를 묶는 CI 스모크 파이프라인 추가 가능

## 10. 기능 추가 작업 시 권장 순서
1. `update_history.md`에 변경 목표/버전 초안 기록
2. UI 추가 전에 상태/설정 스키마를 먼저 설계
3. 워커 로직 변경 시 취소/강제종료 플로우 먼저 점검
4. `hwp_converter.spec` 영향 여부 확인 (새 의존성/hidden import)
5. 수동 테스트 체크리스트 실행

## 11. 수동 테스트 최소 체크리스트
1. 관리자 권한 실행 여부 확인
2. 파일 모드: 드래그 드롭/추가/제거/전체제거
3. 폴더 모드: 하위폴더 포함 on/off 파일 수 미리보기
4. 동일 형식 입력 건너뜀과 사전 점검 다이얼로그
5. 포맷별 변환 성공/실패 처리
6. overwrite on/off 파일명 충돌 처리
7. 결과 다이얼로그(실패 목록 저장, 산출물 감사 필드가 포함된 CSV/JSON 저장, 폴더 열기)
8. 취소 후 강제 종료 경로
9. 백업 옵션 on/off와 실패 자동 재시도
10. 앱 종료 시 트레이/토스트/워커 정리

## 12. 결론
- 이 프로젝트는 "안정화된 COM 자동화 + 고도화된 PyQt UI" 구조이며, 기능 확장 여지는 충분합니다.
- 다만 핵심 안정성 로직(SaveAs 폴백, COM 초기화, 네이티브 DnD)은 절대적인 보호 대상입니다.
- 가장 효율적인 확장 전략은 "설정/결과/검증/자동화 편의 기능"부터 단계적으로 추가하는 방식입니다.

## 13. 정적 분석 및 인코딩 기준 (2026-03-10)
- `pyrightconfig.json`이 저장소 기본 정적 분석 기준이며, `basic` 모드에서 오류 0건을 유지합니다.
- PyQt6의 `menuBar()`, `statusBar()`, `style()`, `mimeData()` 등 `Optional` 반환값은 직접 체이닝하지 않고 로컬 변수 + `None` 가드로 처리합니다.
- COM 객체는 동적 호출 그대로 두지 않고 최소 `Protocol` 타입을 선언해 Pylance가 `None`/동적 속성 오류를 남기지 않게 유지합니다.
- 텍스트 파일은 `.editorconfig` 기준으로 UTF-8, LF, final newline을 사용합니다.

## 14. v8.7 동기화 기준 (2026-04-27)
- 공식 지원 Python 버전은 3.10 이상입니다.
- 앱 버전과 PyInstaller 산출물 이름은 `8.7` / `HWP변환기_v8.7.exe`입니다.
- `.gitignore`는 `build/`, `dist/`, PyInstaller 중간 산출물, 캐시, `backup/`, COM 스모크/결과 리포트 산출물을 제외합니다.
- 실제 한글 COM 동작은 자동 테스트와 별도로 `HWP_COM_SMOKE_TEST_CHECKLIST.md` 기준 수동 검증이 필요합니다.
- 2026-05-12 기능 리스크 보강 내역은 `update_history.md`와 본 구조 분석 문서에 반영되어 있습니다.
