# HwpMate 프로젝트 구조 정밀 분석 (기능 확장용)

## 1. 분석 범위
- 분석 일자: 2026-02-27
- 보강 일자: 2026-03-10
- 대상 저장소: `d:\twbeatles-repos\HwpMate`
- 분석 목적: "다양한 기능 추가"를 위한 현재 구조, 제약, 확장 포인트 파악

## 2. 참고한 문서
- `README.md`
- `claude.md`
- `gemini.md`
- `update_history.md`
- `hwp_converter.spec`
- `hwptopdf-hwpx_v4.py` (메인 코드)
- `hwptopdf-hwpx v3.py` (레거시 참고)

## 3. 저장소 구조 스냅샷

| 파일 | 라인 수 | 역할 |
|---|---:|---|
| `hwptopdf-hwpx_v4.py` | 3659 | 현재 메인 애플리케이션 (GUI + 변환엔진 + 워커 + DnD + 설정) |
| `hwptopdf-hwpx v3.py` | 845 | 레거시 tkinter 기반 버전 |
| `hwp_converter.spec` | 103 | PyInstaller 빌드 설정 (경량화, `uac_admin=True`) |
| `pyrightconfig.json` | 17 | Pylance/pyright 공용 정적 분석 설정 |
| `.editorconfig` | 7 | UTF-8/LF 편집 규칙 |
| `README.md` | 196 | 사용자 관점 기능/설치/사용법 |
| `claude.md` | 78 | 핵심 로직 보존 지침(변경 주의사항) |
| `gemini.md` | 123 | 유지보수/확장 지침(절대 변경 금지 영역 포함) |
| `update_history.md` | 164 | 버전 이력, 기술적 문제 해결 기록 |

현재 구조는 "단일 대형 파일 중심" 아키텍처이며, 배포 편의성에 최적화되어 있습니다.

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
| `ThemeManager` | 다크/라이트 QSS 전체 스타일 제공 |
| `ToastWidget`, `ToastManager` | 우하단 스택형 알림 (최대 3개) |
| 유틸 함수 (`load_config`, `iter_supported_files` 등) | 설정 저장/로드, 경로 정규화, 파일 스캔, 권한 체크 |
| `FileScanWorker` (`QThread`) | 파일/폴더를 비동기 배치 스캔 |
| `HWPConverter` | COM 연결/문서 열기/SaveAs/정리 담당 |
| `ConversionTask` | 입력/출력/상태/오류를 담는 작업 단위 |
| `ConversionWorker` (`QThread`) | 작업 리스트 순차 변환, 백업, 진행 이벤트 전송 |
| `NativeDropFilter` | 관리자 권한에서도 동작하는 네이티브 DnD 처리 |
| `DropArea` | 파일 모드 입력 UI 컴포넌트 |
| `FormatCard` | 변환 포맷 카드 UI 컴포넌트 |
| `ResultDialog` | 완료 결과/실패 목록/폴더 열기 |
| `MainWindow` | 전체 UI 구성, 상태 관리, 워커 제어, 이벤트 오케스트레이션 |

### 4.3 데이터/상태 모델
- 설정 파일: `%USERPROFILE%\.hwp_converter_config.json`
- 기본 설정 키:
  - `config_version`, `theme`, `mode`, `format`, `include_sub`, `same_location`, `overwrite`
- 추가 저장 키:
  - `folder_path`, `output_path`, `last_folder`, `last_output`
- 런타임 주요 상태:
  - `self.file_list`, `self._file_set`
  - `self.tasks`, `self.worker`, `self.file_scan_worker`
  - `self.is_converting`, `self._scan_mode`, `self._force_kill_pending`

## 5. 실제 동작 플로우

### 5.1 파일 수집 플로우
1. 사용자 입력 (파일 선택, 폴더 선택, 네이티브 드롭)
2. `_start_scan()`으로 `FileScanWorker` 실행
3. 배치별 `_on_scan_batch_found()` 호출
4. `_append_files_batch()`에서 중복 제거 + 테이블 렌더링
5. `_on_scan_finished()`에서 상태 라벨 갱신

### 5.2 변환 플로우
1. `_start_conversion()`
2. `_collect_tasks()`로 모드별 작업 생성
3. 필요 시 `_adjust_output_paths()`로 충돌 회피
4. `ConversionWorker` 시작
5. `ConversionWorker.run()` 내부
   - 워커 스레드 `pythoncom.CoInitialize()`
   - `HWPConverter.initialize()`
   - 파일별 `_create_backup()` 후 `convert_file()`
6. `task_completed` 시그널 수신 후 `ResultDialog` 표시
7. `_on_worker_finished()`에서 UI/시그널/상태 정리

## 6. 기능 추가 시 반드시 지켜야 할 핵심 제약
(출처: `claude.md`, `gemini.md`, 코드 본문)

1. SaveAs 이중 전략 유지
- 2-파라미터 실패 시 3-파라미터(`""`) 재시도 로직 필수.

2. 보안/팝업 제어 유지
- `RegisterModule("FilePathCheckDLL", ...)`
- `SetMessageBoxMode(0x00000001)`

3. 스레드 COM 초기화 유지
- 워커 스레드 진입 시 `pythoncom.CoInitialize()` 필수.

4. 자동 백업 로직 유지
- 백업 실패가 전체 변환 실패로 이어지지 않게 현재 방식을 유지.

5. 관리자 권한 + 네이티브 DnD 경로 유지
- UIPI 우회용 WM_DROPFILES 처리 삭제 금지.

6. 배포 제약 고려
- `hwp_converter.spec`의 `hiddenimports`, `uac_admin=True`, excludes 전략을 손상시키지 말 것.

## 7. 확장성 평가

### 7.1 강점
- 비동기 스캔/변환 스레드 분리 완료.
- 변환 포맷 메타데이터(`FORMAT_TYPES`) 중심 확장 구조.
- UI/상태/변환 흐름이 비교적 명확한 메서드 단위로 분리.
- 변환 안정성 관련 폴백/예외 처리 경험치가 코드에 축적됨.

### 7.2 한계
- 단일 파일 3659라인으로 변경 영향 범위가 큼.
- UI 로직과 도메인 로직이 `MainWindow`에 집중됨.
- 자동 테스트 코드 부재.
- 설정 스키마 마이그레이션은 단순 병합 방식(복잡한 호환 로직 없음).
- 워커/스캐너 상태 관리가 객체 필드 기반이라 기능이 늘수록 복잡도 상승 가능.

## 8. 추천 기능 추가 항목 (우선순위)

### 8.1 단기(낮은 리스크)
1. 변환 프리셋 저장/불러오기
- 예: "PDF-사내공유", "DOCX-검수용"
- 영향 영역: 설정 로드/저장, UI(프리셋 콤보), `_save_settings`, `_start_conversion`

2. 실패 자동 재시도 옵션
- 예: 실패 파일 N회 재시도, 대기시간 설정
- 영향 영역: `ConversionWorker.run()`, 결과 다이얼로그

3. 결과 리포트 확장(CSV/JSON)
- 현재 txt 실패 목록 외 전체 결과 리포트
- 영향 영역: `ResultDialog`, `_on_task_completed`

4. 출력 파일명 템플릿
- 예: `{name}_{date}` `{name}_{format}`
- 영향 영역: `_collect_tasks`, `_adjust_output_paths`

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
`claude.md`의 "단일 파일 유지" 원칙을 깨지 않고도 아래 개선은 가능:

1. 파일 내부 섹션화 강화
- `# region` 스타일 주석으로 UI/Worker/Converter 경계 명확화

2. 데이터 클래스 도입
- `ConversionTask`를 `@dataclass`로 전환해 필드 명시성 강화

3. 설정 접근 래퍼 함수 추가
- `get_config_*`, `set_config_*` 함수로 키 분산 접근 축소

4. 상태 전이 명시화
- `is_converting`, `_scan_mode`, `_force_kill_pending` 전이를 표로 문서화 및 가드 함수 통일

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
4. 포맷별 변환 성공/실패 처리
5. overwrite on/off 파일명 충돌 처리
6. 취소 후 강제 종료 경로
7. 결과 다이얼로그(실패 목록 저장/폴더 열기)
8. 앱 종료 시 트레이/토스트/워커 정리

## 12. 결론
- 이 프로젝트는 "안정화된 COM 자동화 + 고도화된 PyQt UI" 구조이며, 기능 확장 여지는 충분합니다.
- 다만 핵심 안정성 로직(SaveAs 폴백, COM 초기화, 네이티브 DnD)은 절대적인 보호 대상입니다.
- 가장 효율적인 확장 전략은 "설정/결과/검증/자동화 편의 기능"부터 단계적으로 추가하는 방식입니다.

## 13. 정적 분석 및 인코딩 기준 (2026-03-10)
- `pyrightconfig.json`이 저장소 기본 정적 분석 기준이며, `basic` 모드에서 오류 0건을 유지합니다.
- PyQt6의 `menuBar()`, `statusBar()`, `style()`, `mimeData()` 등 `Optional` 반환값은 직접 체이닝하지 않고 로컬 변수 + `None` 가드로 처리합니다.
- COM 객체는 동적 호출 그대로 두지 않고 최소 `Protocol` 타입을 선언해 Pylance가 `None`/동적 속성 오류를 남기지 않게 유지합니다.
- 텍스트 파일은 `.editorconfig` 기준으로 UTF-8, LF, final newline을 사용합니다.
