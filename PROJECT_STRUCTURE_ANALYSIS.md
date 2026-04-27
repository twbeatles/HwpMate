# HwpMate 프로젝트 구조 정밀 분석 (기능 확장용)

## 1. 분석 범위
- 분석 일자: 2026-02-27
- 보강 일자: 2026-03-10
- 추가 보강 일자: 2026-03-18
- v8.7 반영 일자: 2026-04-27
- 대상 저장소: `c:\twbeatles-repos\HwpMate`
- 분석 목적: "다양한 기능 추가"를 위한 현재 구조, 제약, 확장 포인트 파악

## 2. 참고한 문서
- `README.md`
- `claude.md`
- `gemini.md`
- `update_history.md`
- `IMPLEMENTATION_RISK_REVIEW.md`
- `HWP_COM_SMOKE_TEST_CHECKLIST.md`
- `hwp_converter.spec`
- `hwptopdf-hwpx_v4.py` (루트 엔트리포인트 래퍼)
- `hwpmate/` (현행 메인 코드)
- `legacy/hwptopdf-hwpx v3.py` (레거시 참고)

## 3. 저장소 구조 스냅샷

| 파일 | 라인 수 | 역할 |
|---|---:|---|
| `hwptopdf-hwpx_v4.py` | 5 | 패키지 진입용 얇은 래퍼 |
| `hwpmate/` | 모듈 분리 | 현재 메인 애플리케이션 (GUI + 변환엔진 + 워커 + DnD + 설정) |
| `legacy/hwptopdf-hwpx v3.py` | 828 | 레거시 tkinter 기반 버전 |
| `hwp_converter.spec` | 107 | PyInstaller 빌드 설정 (경량화, `uac_admin=True`) |
| `pyrightconfig.json` | 15 | Pylance/pyright 공용 정적 분석 설정 |
| `.editorconfig` | 13 | UTF-8/LF 편집 규칙 |
| `README.md` | 131 | 사용자 관점 기능/설치/사용법 |
| `IMPLEMENTATION_RISK_REVIEW.md` | 46 | v8.7 구현 리스크 개선 완료 보고서 |
| `HWP_COM_SMOKE_TEST_CHECKLIST.md` | 48 | 실제 한글 COM 수동 검증 체크리스트 |
| `claude.md` | 110 | 핵심 로직 보존 지침(변경 주의사항) |
| `gemini.md` | 91 | 유지보수/확장 지침(절대 변경 금지 영역 포함) |
| `update_history.md` | 123 | 버전 이력, 기술적 문제 해결 기록 |

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
| `ui/main_window.py` | UI 상태 관리, 모드별 DnD 분기, 종료 제어 오케스트레이션 |
| `ui/main_window_ui.py` | 메인 윈도우 레이아웃 빌더 |

### 4.3 데이터/상태 모델
- 설정 파일: `%USERPROFILE%\.hwp_converter_config.json`
- 기본 설정 키:
  - `config_version`, `theme`, `mode`, `format`, `include_sub`, `same_location`, `overwrite`
  - `backup_enabled`, `retry_count`
- 추가 저장 키:
  - `folder_path`, `output_path`, `last_folder`, `last_output`
- 런타임 주요 상태:
  - `self.file_list`, `self._file_set`
  - `self.plan`, `self.tasks`, `self.last_summary`, `self.worker`, `self.file_scan_worker`
  - `self.is_converting`, `self._scan_mode`, `self._force_kill_pending`, `self._close_after_worker`

## 5. 실제 동작 플로우

### 5.1 파일 수집 플로우
1. 사용자 입력 (파일 선택, 폴더 선택, 네이티브 드롭)
2. `_start_scan()`으로 `FileScanWorker` 실행
3. 배치별 `_on_scan_batch_found()` 호출
4. `_append_files_batch()`에서 중복 제거 + 테이블 렌더링
5. `_on_scan_finished()`에서 상태 라벨 갱신

### 5.2 변환 플로우
1. `_start_conversion()`
2. `_collect_tasks()`로 `PlannedConversion` 생성
3. 필요 시 `_adjust_output_paths()`로 충돌 회피 수 계산
4. `PreflightDialog`로 실행 대상/건너뜀/경고 확인
5. 실행 대상 없이 동일 형식 건너뜀만 있으면 즉시 `ResultDialog` 표시
6. 실행 대상이 있으면 `ConversionWorker` 시작
7. `ConversionWorker.run()` 내부
   - 워커 스레드 `pythoncom.CoInitialize()`
   - `HWPConverter.initialize()`
   - 파일별 선택적 `_create_backup()` 후 `convert_file()`
   - 실패 시 설정된 횟수만큼 재시도
   - 취소 시 남은 task를 `취소됨`으로 마킹
   - `ConversionSummary` 생성
8. `task_completed` 시그널 수신 후 `ResultDialog` 표시
9. 필요 시 실패 TXT / 결과 CSV·JSON 저장
10. `_on_worker_finished()`에서 UI/시그널/상태 정리 및 종료 대기 처리

## 6. 기능 추가 시 반드시 지켜야 할 핵심 제약
(출처: `claude.md`, `gemini.md`, 코드 본문)

1. SaveAs 이중 전략 유지
- 2-파라미터 실패 시 3-파라미터(`""`) 재시도 로직 필수.
- `Open()`/`SaveAs()`가 명시적으로 `False`를 반환하면 실패로 처리.
- 출력 파일 존재 및 0바이트 초과 검증 유지.

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

7. 강제 종료 범위 제한 유지
- 시스템 전체 한글 프로세스 종료로 되돌리지 말고, 앱이 추적한 PID만 종료.

8. 배포 제약 고려
- `hwp_converter.spec`의 `hiddenimports`, `uac_admin=True`, excludes 전략을 손상시키지 말 것.

## 7. 확장성 평가

### 7.1 강점
- 비동기 스캔/변환 스레드 분리 완료.
- 변환 포맷 메타데이터(`FORMAT_TYPES`) 중심 확장 구조.
- UI/상태/변환 흐름이 비교적 명확한 메서드 단위로 분리.
- 변환 안정성 관련 폴백/예외 처리 경험치가 코드에 축적됨.

### 7.2 한계
- `MainWindow`가 여전히 가장 큰 조정 지점이라 UI 상태 전이가 이곳에 비교적 많이 남아 있습니다.
- PyQt 위젯 생성은 분리됐지만, 런타임 오케스트레이션은 단일 클래스 중심입니다.
- 자동 테스트는 순수 로직 계층 위주이며, GUI/COM 경로는 여전히 수동 검증 비중이 높습니다.
- 설정 스키마 마이그레이션은 기본값 병합 방식이며, v8.7에서 `config_version=2`입니다.
- 워커/스캐너 상태 관리가 객체 필드 기반이라 기능이 늘수록 복잡도 상승 가능.

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

1. `MainWindow` 상태 전이 축소
- 변환 상태, 스캔 상태, 토스트/트레이 상태를 별도 상태 객체로 더 분리 가능

2. UI 액션 서비스화
- 폴더 열기, 실패 목록 내보내기, 사용자 알림 정책을 더 작은 서비스로 분리 가능

3. GUI/COM 검증 보강
- Qt 상호작용 테스트는 추가됐으므로, 실제 HWP COM이 설치된 환경에서 도는 스모크 테스트/수동 점검 체계를 더 강화할 수 있음

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
7. 결과 다이얼로그(실패 목록 저장, CSV/JSON 저장, 폴더 열기)
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
- `.gitignore`는 `build/`, `dist/`, PyInstaller 중간 산출물, 캐시, `backup/` 폴더를 제외합니다.
- 실제 한글 COM 동작은 자동 테스트와 별도로 `HWP_COM_SMOKE_TEST_CHECKLIST.md` 기준 수동 검증이 필요합니다.
- `IMPLEMENTATION_RISK_REVIEW.md`는 v8.7에서 반영된 리스크 개선 내역을 요약합니다.
