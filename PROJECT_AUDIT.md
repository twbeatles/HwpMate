# Project Audit

- 감사 일자: 2026-09-15
- 대상 커밋: `a63aa1a` (main, 앱 버전 `9.1.0`)
- 감사 범위: `hwptopdf-hwpx_v4.py` → `hwpmate/` 패키지, `scripts/`, `tools/`, `tests/`, 빌드/릴리즈 설정 (`legacy/`는 제외)
- 원칙: 코드 수정 없음. 반증되지 않은 문제만 이슈로 기록. 재현 스크립트는 저장소 밖 임시 디렉터리에서만 실행.

> **조치 현황 (2026-09-15 갱신):** ISSUE-001~007과 §5의 Confirmed/Likely Gap을 모두 수정했습니다. 이후 실제 한글 2022 COM 점검에서 추가 결함 5건(ISSUE-008~012)을 발견해 함께 수정했습니다. 상세 결과와 검증은 **§10 조치 결과**를 참고하세요. 아래 §1~§9는 감사 당시 기록을 그대로 유지합니다.

---

## 1. Executive Summary

### 전체 상태

HwpMate는 COM 경계(소유 PID 한정 강제 종료, SaveAs 2→3 인자 폴백, `%PDF` 매직 검증, 산출물 스냅샷 기반 성공 판정), 설정/결과 원자 저장, 서명 기반 업데이트 매니페스트 검증 등 **핵심 안전장치가 잘 갖춰진 편**입니다. `pytest` 174개 통과, `pyright` 0 errors 상태입니다.

다만 이번 감사에서 **기존 테스트가 잡지 못하는 실제 결함**을 재현으로 확인했습니다. 대부분 "현실적인 입력 조건"(느린 폴더 스캔, 원본과 같은 stem의 파일, 반복 백업, 콘솔 없는 헬퍼 프로세스)에서만 드러나며, 테스트가 이 조건을 모킹하거나 다루지 않습니다.

### 전체 위험도: **Medium**

### 가장 중요한 문제

| # | 문제 | 우선순위 | 신뢰도 |
|---|------|---------|--------|
| ISSUE-001 | 폴더 스캔이 2초 이상 걸리는 폴더는 **폴더 모드 변환을 영구히 시작할 수 없음** | High | Confirmed (재현) |
| ISSUE-002 | 「덮어쓰기」 + HWP/HWPX 대상 변환 시 **같은 stem의 원본 문서를 덮어쓰며 백업도 없음** | High | Confirmed (계획 단계 재현) |
| ISSUE-003 | 백업 보관 개수 정리가 **다른 파일의 백업까지 삭제** (`report.hwp` 정리 시 `report_1.hwp` 백업 삭제) | Medium | Confirmed (재현) |
| ISSUE-004 | CLI가 README에 문서화된 `--retry`와 기본 백업을 **실제로 적용하지 않음**, 빈 폴더 시 미처리 예외 | Medium | Confirmed (재현) |
| ISSUE-005 | 업데이트 헬퍼의 부모 종료 대기가 **무력화**(`os.kill(pid, 0)`), 실패 시 "롤백됨"으로 오보고되고 이후 같은 버전 업데이트가 계속 실패 | Medium | Confirmed(대기 무력화) / Likely(경합 실패) |

### 데이터 손상/유실 가능성

**있음 (조건부).**
- ISSUE-002: 사용자가 덮어쓰기를 켜고 HWP↔HWPX 변환 시, 같은 폴더의 동일 stem 원본(`a.hwp`/`a.hwpx`)이 변환 결과로 덮어써집니다. 백업은 변환 입력만 복사하므로 덮어써진 원본은 복구할 수 없습니다.
- ISSUE-003: 원본이 아닌 **백업 사본**이 의도보다 많이 삭제됩니다. 원본 자체는 변환으로 수정되지 않으므로 즉시 유실은 아니지만, 백업 이력이라는 안전장치가 약해집니다.

### 가장 먼저 수정해야 할 영역

1. 폴더 모드 변환 시작 흐름의 스캔 대기(`ConversionController._ensure_folder_scan_ready`)
2. 출력 경로 할당에서 "입력/건너뜀 원본과의 충돌" 차단(`TaskPlanner.allocate_output_path`)
3. 백업 정리의 stem 매칭(`_prune_old_backups`)

---

## 2. Project Understanding

### 프로젝트 목적

한컴오피스 한글이 설치된 Windows 10/11에서 HWP/HWPX 문서를 PDF, DOCX, HWP/HWPX, ODT, HTML, RTF, TXT, PNG/JPG/BMP/GIF로 **대량 일괄 변환**하는 PyQt6 GUI + 헤드리스 CLI 도구입니다. 한글 COM 자동화(`HWPControl.HwpCtrl.1` 등)를 사용하며 관리자 권한 실행을 전제합니다.

### 주요 Entrypoint

| Entrypoint | 위치 | 역할 |
|-----------|------|------|
| GUI | `hwptopdf-hwpx_v4.py` → `hwpmate/bootstrap.py` → `hwpmate/app.py:main` | pywin32/관리자 확인 → `SingleInstanceLock` → `MainWindow` |
| CLI 변환 | `hwpmate/app.py:_run_cli_conversion` (`--input`) | GUI 없이 계획 → 변환 루프 |
| 스모크 | `hwpmate/app.py:_run_smoke` (`--smoke`) | import 무결성 JSON 출력 |
| 업데이트 헬퍼 | `hwpmate/app.py:_run_apply_update` (`--apply-update`) | 스테이징 exe 교체·스모크·롤백 |
| 보조 스크립트 | `scripts/apply_update.py`, `scripts/build_update_manifest.py`, `tools/hwp_com_smoke.py` | 릴리즈/수동 검증용 |

### 핵심 모듈

- `services/task_planner.py` — 작업 계획, 동일 형식 건너뜀, 출력 충돌 회피(`allocate_output_path`)
- `services/artifact_policy.py` — 이미지/HTML 보조 산출물 stem 매칭
- `services/hwp_converter/converter.py` — COM 초기화, 보안 모듈 등록, `convert_file`(Open → 인쇄 리셋 → SaveAs/PrintToPDF → 산출물 검증), 소유 PID 강제 종료
- `services/hwp_print_settings/` — PrintMethod 리셋, PrintToPDFEx/RunToPDF (물리 Print 미사용 확인)
- `services/hwp_security_module.py`, `services/hwp_security_session.py` — DLL 설치·SHA-256·HKCU 등록, 자동 클릭 정책
- `workers/conversion_worker/` — QThread 변환 루프, 백업, 재시도, 200건 재순환, 요약
- `workers/file_scan_worker.py` — 비동기 폴더/파일 스캔
- `ui/main_window_controllers/` — conversion / file_selection / lifecycle / appearance / native_drop / update
- `services/update_manifest.py`, `services/update_installer.py` — Ed25519 매니페스트 검증, 스트리밍 다운로드, 교체/롤백
- `config_repository.py`, `ui/dialogs/atomic_io.py` — 설정/결과 원자 저장

### 데이터 저장 방식 (DB 없음, 전부 파일)

| 데이터 | 위치 | 쓰기 방식 |
|--------|------|----------|
| 설정 | `~/.hwp_converter_config.json` | 임시 파일 → `replace` (원자적), 손상 시 `.bak`로 이동 |
| 원본 백업 | 입력 파일 옆 `backup/` | `shutil.copy2` + stem별 보관 개수 정리 |
| 변환 산출물 | 입력 옆 또는 사용자 출력 폴더 | 한글 COM SaveAs/PrintToPDF |
| 결과 CSV/JSON/TXT | 사용자 선택 경로 | 임시 파일 → `replace` |
| 로그 | `~/.hwp_converter/logs` 등 | RotatingFileHandler |
| 보안 DLL | `%LOCALAPPDATA%\HwpMate\security` + HKCU 레지스트리 | `.tmp` 복사 후 교체, SHA-256 검증 |
| 업데이트 스테이징/결과 | `%LOCALAPPDATA%\HwpMate\updates` | `open("xb")`, 결과 JSON 원자 저장 |
| 단일 인스턴스 잠금 | `%LOCALAPPDATA%\HwpMate\HwpMate.lock` | `QLockFile` |

### 외부 의존성

- 한컴오피스 한글 COM (런타임 필수), pywin32, PyQt6, cryptography(Ed25519), GitHub raw/Releases (HTTPS)
- 빌드: PyInstaller onefile(`console=False`, `uac_admin=True`), CI: GitHub Actions(windows-latest, Python 3.12)

### 핵심 실행 흐름

**GUI 폴더 변환**

`변환 시작(Ctrl+Enter)` → `ConversionController.start_conversion` → (`is_planning` 잠금) → `_ensure_folder_scan_ready`(비동기 재스캔 + 대기) → `validate_output_settings` → `collect_tasks` → `TaskPlanner.build_tasks`(동일 형식 → `skipped_tasks`) → `resolve_output_conflicts` → `PreflightDialog` → `ConversionWorker.run`(CoInitialize) → `HWPConverter.initialize(manage_com_apartment=False)` → 작업별 [`create_backup` → 런타임 `allocate_output_path` → `convert_file` × (retry+1)] → 200건마다 재순환 → `ConversionSummary` → `ResultDialog` → CSV/JSON 원자 저장

**CLI 변환**

`--input` → `_run_cli_conversion` → `build_tasks` → `resolve_output_conflicts` → `HWPConverter.initialize()` → 작업별 `convert_file` 1회 → 요약 출력 → exit code

**자동 업데이트**

`showEvent` → 3초 후 `UpdateCheckWorker`(매니페스트 다운로드 → 서명 검증 → 버전 비교) → `UpdateDialog` → `UpdateDownloadWorker`(스트리밍 + SHA-256/크기) → `_apply_and_restart` → 헬퍼 exe 복사·실행 → 부모 종료 → 헬퍼 `_wait_for_parent` → `apply_staged_update`(백업 복사 → `os.replace` → `--smoke` → 실패 시 롤백) → 결과 JSON → 다음 실행 시 토스트

---

## 3. Audit Coverage & Limitations

### 실제 확인한 주요 모듈 (전체 소스 열람)

`app.py`, `app_instance.py`, `models.py`, `constants.py`, `config_repository.py`, `path_utils.py`, `logging_config.py`, `services/task_planner.py`, `services/artifact_policy.py`, `services/file_selection_store.py`, `services/hwp_converter/*`, `services/hwp_print_settings/{pdf_export,print_reset,printers}.py`, `services/hwp_security_module.py`, `services/hwp_security_session.py`, `services/update_installer.py`, `services/update_manifest.py`, `workers/conversion_worker/*`, `workers/file_scan_worker.py`, `ui/main_window.py`, `ui/main_window_controllers/{conversion,file_selection,lifecycle,appearance,native_drop,update}`, `ui/dialogs/{preflight,result,atomic_io,update_dialog}.py`, `windows_integration/{admin,hwp_window_control,native_drop}.py`, `scripts/apply_update.py`, `hwp_converter.spec`(주요 옵션), `.github/workflows/release.yml`.

일부만 확인: `ui/main_window_ui/builder.py`(버튼 연결부), `windows_integration/window_query.py`, 테마/토스트/위젯(기능 영향 낮음).

### CodeGraph로 분석한 호출 관계

- 프롬프트 시점 CodeGraph 컨텍스트와 `codegraph_explore`로 다음 심볼의 caller/blast radius 확인:
  - `resolve_output_conflicts` — 호출자 7곳 (`app.py` CLI, `ConversionController.adjust_output_paths` 경유 GUI 시작/실패 재변환), 테스트 `tests/test_task_planner.py`
  - `allocate_output_path` — `TaskPlanner.resolve_output_conflicts` + `ConversionWorker.run`(런타임 재할당)
  - `create_backup` / `_prune_old_backups` — `ConversionWorker._create_backup` 경유 단일 경로 (CLI에서는 호출 없음)
  - `refresh_folder_scan_for_conversion` — `ConversionController._ensure_folder_scan_ready` 단일 호출자
  - `_wait_for_parent` — `hwpmate/app.py`와 `scripts/apply_update.py`에 동일 구현 2개
- 그 외 흐름(업데이트 컨트롤러, 수명주기, 워커 루프)은 파일 직접 열람으로 추적했습니다.

### 실행한 테스트/검증

| 명령 | 결과 |
|------|------|
| `python -m pytest -q` | **174 passed** (5.79s, 로컬 Python 3.14.7) |
| `pyright .` | **0 errors, 0 warnings** (pyright 1.1.411) |
| 재현: 이미지/HTML 출력 충돌 (`TaskPlanner`, 임시 폴더) | PNG/HTML 모두 `report (1).*`로 변경됨 확인 |
| 재현: 덮어쓰기 + HWP 대상 계획 | `a.hwpx → a.hwp` 작업, `a.hwp`는 건너뜀, 이름 변경 0 확인 |
| 재현: 백업 정리 (`create_backup`, max_files=3) | `report.hwp` 백업 시 `report_0~2.hwp` 백업 삭제 확인 |
| 재현: 느린 폴더 스캔(3초) + `start_conversion` (offscreen Qt, 가짜 설정) | 3회 연속 사전 점검 도달 실패, 동일 조건 0.2초 스캔은 3회 모두 도달 |
| 재현: CLI (가짜 컨버터) | `--retry 3`에도 변환 1회, 백업 0회, 건너뜀 0건 출력, 빈 폴더 `ValueError` 미처리 |
| 검증: `os.kill(pid, 0)` (자체 생성 자식 프로세스) | 콘솔 없는 호출자(`pythonw` + `CREATE_NO_WINDOW`)에서 대상이 살아 있는데도 즉시 `OSError(WinError 6)` |

### 확인하지 못한 환경/외부 서비스

- **한컴오피스 한글 COM 실제 변환 미실행** (`tools/hwp_com_smoke.py` 미실행). SaveAs의 기존 파일 덮어쓰기 동작, 이미지 다중 페이지 산출물 명명 규칙, PrintToPDFEx 실제 동작은 코드로만 판단했습니다.
- PyInstaller 빌드 및 frozen exe 업데이트 교체 E2E 미실행 (onefile 부트로더의 exe 잠금 해제 시점 미측정).
- GitHub 매니페스트/릴리즈 네트워크 호출 미실행.
- 관리자 권한 네이티브 드래그 앤 드롭, 실제 트레이/최소화 복원 이벤트 미확인.
- CI의 Python 3.12가 아닌 로컬 3.14.7로 테스트 실행.

### 반증되어 이슈에서 제외한 후보

- **다운로드 소켓 읽기 무한 대기** (이전 감사 지적): `urlopen(timeout=30)`은 소켓 타임아웃으로 이후 `read()`에도 적용되므로 제외.
- **업데이트 다운로드 취소 미반영** (이전 감사 지적): `cancel_check`가 스트림/스테이징 모두에 연결되어 있고 스테이징 파일 삭제까지 구현되어 제외.
- **대량 변환 COM 재순환 부재** (이전 감사 지적): `CONVERTER_RECYCLE_BATCH_COUNT=200`으로 구현되어 제외.
- **강제 종료가 사용자 한글 세션까지 종료**: `kill_owned_processes`가 소유 PID + 실시간 이미지명 재확인으로 제한되어 제외.
- **물리 프린터 출력**: `CreateAction("Print")`는 `GetDefault`/`SetItem`만 호출하고 `Execute`는 없음을 확인해 제외.
- **설정/결과 파일 부분 쓰기 손상**: 임시 파일 + `replace`로 보호되어 제외.
- **매니페스트 위조/다운그레이드**: 서명 검증 → 버전 비교 → HTTPS → 해시/크기 → 만료 순으로 검증되어 제외.

### 분석상의 한계

- CodeGraph는 Qt 시그널/람다 연결(`worker.finished.connect(...)`)의 동적 호출을 완전히 따라가지 못해, 해당 부분은 직접 열람으로 보완했습니다.
- PyQt6에서 슬롯 내부 `sys.exit(0)` 처리 방식은 버전별로 달라 ISSUE-006은 Likely로 분류했습니다.

---

## 4. High-Risk Issues

### [ISSUE-001] 폴더 스캔이 2초를 넘는 폴더는 폴더 모드 변환을 시작할 수 없음

- **위치:** `hwpmate/ui/main_window_controllers/conversion/controller.py` / `ConversionController._ensure_folder_scan_ready` (221–250행), `hwpmate/ui/main_window_controllers/file_selection/controller.py` / `FileSelectionController.refresh_folder_scan_for_conversion` (150–154행), `hwpmate/constants.py:52` (`SCAN_CANCEL_WAIT_MS = 2000`)
- **우선순위:** High
- **신뢰도:** Confirmed
- **문제:** 변환 시작 시 `_ensure_folder_scan_ready`는 캐시를 무효화하고 새 스캔을 시작한 뒤 `wait_for_active_scan(SCAN_CANCEL_WAIT_MS)`로 **2초만** 기다립니다. 2초 안에 끝나지 않으면 `ValueError("폴더 스캔이 아직 종료되지 않았습니다…")`로 중단합니다. 사용자가 다시 시작하면 `start_conversion`이 진행 중인 스캔을 최대 120초 기다려 캐시를 얻지만, 곧이어 `_ensure_folder_scan_ready`가 **그 캐시를 다시 무효화하고 새 스캔을 시작해 또 2초만 대기**합니다. 따라서 스캔 시간이 2초를 넘는 폴더는 몇 번을 눌러도 사전 점검까지 가지 못합니다.
- **발생 조건:** 폴더 모드 + 스캔에 2초 이상 걸리는 폴더 (하위 폴더 포함 대량 파일, 네트워크 드라이브/NAS, 콜드 캐시 HDD, 백신 실시간 검사 등).
- **영향:** 제품의 핵심 사용 사례(대량 폴더 일괄 변환)가 대형/원격 폴더에서 **전면 불가**. 우회책은 파일 모드로 전환해 파일을 직접 추가하는 것뿐입니다.
- **근거:** 오프스크린 Qt 재현에서 `FileScanWorker` 스캔을 3초로 지연시키자 `start_conversion` 3회 모두 사전 점검에 도달하지 못하고 동일 경고가 발생했습니다. 같은 스크립트에서 0.2초로 바꾸면 3회 모두 도달했습니다.
- **반증 확인:**
  - `start_conversion`의 120초 대기(`FOLDER_SCAN_WAIT_MS`)는 **이미 진행 중인** 스캔만 기다리며, 직후 `_ensure_folder_scan_ready`가 캐시를 버리므로 보호가 되지 않습니다.
  - `_ensure_folder_scan_ready` 하단의 `wait_for_active_scan(FOLDER_SCAN_WAIT_MS)`는 앞의 `refresh…`가 False를 반환하면 `raise`로 먼저 빠져나가 도달하지 못합니다.
  - 기존 테스트(`test_collect_tasks_requires_folder_cache_after_wait` 등)는 스캔이 즉시 끝나거나 워커를 모킹해 이 경로를 검증하지 않습니다.
- **호출/영향 범위:** `start_conversion` → `_ensure_folder_scan_ready` → `refresh_folder_scan_for_conversion` → `start_folder_preview_scan` → `wait_for_active_scan`. GUI 폴더 모드 변환 전체. CLI는 영향 없음.
- **권장 수정 방향:** `refresh_folder_scan_for_conversion`은 스캔 시작만 하고 대기는 `_ensure_folder_scan_ready` 하단의 `FOLDER_SCAN_WAIT_MS` 대기로 일원화하세요. 또는 직전에 완료된 신선한 캐시(`FOLDER_SCAN_CACHE_CONVERT_MAX_AGE_SECONDS` 이내)가 있으면 재스캔을 생략하세요. 이때 `refresh`가 `start_scan` 내부의 `_input_locked()`를 거치며 "작업 준비 중에는 입력을 변경할 수 없습니다" 토스트를 띄우는 부작용도 함께 정리하는 것이 좋습니다.
- **필요한 회귀 테스트:** `FileScanWorker` 스캔을 3초 지연(`iter_supported_files` 패치)시킨 뒤 폴더 모드 `start_conversion` 1회 호출 시 `_confirm_preflight_and_start_worker`가 호출되고 경고 박스가 뜨지 않아야 합니다.

---

### [ISSUE-002] 「기존 파일 덮어쓰기」 + HWP/HWPX 대상 변환이 같은 stem의 원본 문서를 덮어씀

- **위치:** `hwpmate/services/task_planner.py` / `TaskPlanner.allocate_output_path` (210–263행), `build_tasks` (85–95, 144–155행 — 동일 형식 입력을 `skipped_tasks`로만 분리), `hwpmate/workers/conversion_worker/worker.py:138-143` (런타임 재할당)
- **우선순위:** High
- **신뢰도:** Confirmed (계획 단계 재현) / 실제 COM SaveAs 덮어쓰기는 Likely (한글 미실행)
- **문제:** 출력 경로 충돌 검사는 (a) 같은 배치의 **실행 작업 출력끼리**의 중복과 (b) `overwrite=False`일 때만 기존 파일 존재를 봅니다. **입력 파일, 특히 동일 형식이라 건너뛴 원본은 충돌 대상에 포함되지 않습니다.** 덮어쓰기를 켜고 HWP로 변환하면 `a.hwpx → a.hwp` 작업이 만들어지고, 같은 폴더의 원본 `a.hwp`(건너뜀 처리)가 그대로 출력 경로가 됩니다.
- **발생 조건:** `overwrite=True` + 대상 형식 `HWP` 또는 `HWPX` + 같은 출력 폴더에 동일 stem의 `.hwp`/`.hwpx` 원본 공존 (예: 과거 HWP→HWPX 일괄 변환 결과가 원본 옆에 남은 폴더). GUI·CLI(`--overwrite`) 모두 해당합니다.
- **영향:** 원본 한글 문서가 다른 파일의 변환 결과로 **교체되어 유실**됩니다. 자동 백업은 변환 입력(`a.hwpx`)만 복사하므로 덮어써진 `a.hwp`는 복구 수단이 없습니다. 두 파일 내용이 다르면(편집 이력이 있으면) 실제 데이터 손실입니다.
- **근거:** 임시 폴더에 `a.hwp`, `a.hwpx`를 만들고 `build_tasks(format_type="HWP", same_location=True, overwrite=True)` + `resolve_output_conflicts(overwrite=True)`를 실행한 결과, 작업 `[('a.hwpx', 'a.hwp')]`, 건너뜀 `['a.hwp']`, 이름 변경 0건이었습니다.
- **반증 확인:**
  - `PreflightDialog._blocking_errors`는 입력 존재/읽기, 경로 길이, 출력 폴더 쓰기 권한만 검사하며 출력=원본 충돌은 검사하지 않습니다.
  - 워커 런타임 `allocate_output_path`도 같은 함수를 `overwrite=True`로 호출하므로 막지 못합니다.
  - `remove_new_attempt_artifacts`는 사전에 존재한 경로를 지우지 않지만, 이는 실패 정리용일 뿐 성공 SaveAs의 덮어쓰기를 막지 않습니다.
  - UI 문구는 "기존 파일 덮어쓰기"로, 사용자가 이전 **변환 결과물** 덮어쓰기를 기대할 수는 있어도 같은 배치의 원본 입력 덮어쓰기까지 동의했다고 보기 어렵습니다.
- **호출/영향 범위:** `resolve_output_conflicts` 호출자 7곳(CLI `_run_cli_conversion`, GUI `start_conversion`/`retry_failed_tasks`의 `adjust_output_paths`) + `ConversionWorker.run`의 런타임 재할당.
- **권장 수정 방향:** `allocate_output_path`에서 `overwrite` 여부와 무관하게 **출력 경로가 지원 확장자(`.hwp`/`.hwpx`) 원본 문서이거나 계획의 입력/건너뜀 경로 집합에 속하면** 번호를 붙이세요. 최소한 사전 점검에서 차단 오류로 표시해야 합니다. 덮어쓰기 대상 기존 파일을 백업하는 선택지도 고려할 수 있습니다.
- **필요한 회귀 테스트:**
  - 입력 `a.hwp`, `a.hwpx`, 대상 `HWP`, `overwrite=True` → 작업 출력이 `a (1).hwp`이거나 계획이 차단 경고를 포함해야 합니다.
  - 대칭 케이스로 대상 `HWPX`에서도 동일합니다.
  - `ConversionWorker` 런타임 재할당에서도 원본 경로가 선택되지 않아야 합니다.

---

### [ISSUE-003] 백업 보관 개수 정리가 다른 파일의 백업을 삭제함

- **위치:** `hwpmate/workers/conversion_worker/backup.py` / `_prune_old_backups` (24–66행, 특히 38·46행 `prefix = f"{stem}_"` / `startswith(prefix)`)
- **우선순위:** Medium
- **신뢰도:** Confirmed
- **문제:** 같은 stem 백업을 고르는 기준이 `이름.startswith(stem + "_")`입니다. 백업 이름 형식이 `{stem}_{timestamp}{suffix}`이므로, `report.hwp`를 정리할 때 `report_1.hwp`, `report_final.hwp`, `report_2024.hwp` 등 **`report_`로 시작하는 모든 파일의 백업**이 후보가 되어 오래된 것부터 삭제됩니다. 게다가 `shutil.copy2`가 원본 mtime을 보존하므로 정렬 기준 `st_mtime`은 백업 시각이 아니라 **원본 파일의 수정 시각**입니다. 따라서 오래전에 수정된 다른 문서의 백업이 방금 만들어졌어도 먼저 삭제됩니다.
- **발생 조건:** 같은 폴더에 `공문.hwp`와 `공문_001.hwp … 공문_050.hwp`처럼 접두사를 공유하는 파일이 있고, 백업 사용(기본값) 상태로 두 번 이상 변환하거나 `공문.hwp`만 따로 변환할 때. 공문서 번호 체계에서 흔한 명명입니다.
- **영향:** 사용자가 설정한 "stem별 최대 보관 개수"보다 훨씬 적은 백업만 남습니다. 파일 모드로 `공문.hwp`만 변환하면 다른 파일들의 **유일한 백업**이 지워질 수 있습니다. 원본은 변환으로 수정되지 않아 즉시 유실은 아니지만, 원본이 이후 손상/수정됐을 때 복구 이력이 사라집니다.
- **근거:** 임시 폴더에서 `report_0~4.hwp` 백업 5개 생성 후 `create_backup(report.hwp, max_files=3)`을 실행하자 `report_0`, `report_1`, `report_2`의 백업이 삭제됐습니다.
- **반증 확인:** `keep_path` 보호는 방금 만든 백업 1개만 보호합니다. `test_create_backup_prunes_old_stem_files`는 단일 stem만 사용해 교차 stem 케이스를 검증하지 않습니다. 워커가 백업 실패를 무시하는 정책(CLAUDE.md)은 삭제 오동작을 막지 않습니다.
- **호출/영향 범위:** `ConversionWorker.run` → `_create_backup` → `create_backup` → `_prune_old_backups`. GUI 변환·실패 재변환 모두 해당(CLI는 백업 자체가 없음, ISSUE-004).
- **권장 수정 방향:** 후보 판정을 `^{re.escape(stem)}_\d{8}_\d{6}_\d{6}(_\d+)?{suffix}$` 형태의 정확한 패턴 매칭(대소문자 무시)으로 바꾸고, 정렬은 파일명의 타임스탬프를 기준으로 하세요. 백업 파일에 원본 mtime 보존이 필요 없다면 `copy2` 대신 생성 시각을 쓰는 방법도 있습니다.
- **필요한 회귀 테스트:**
  - `report_0~4.hwp` 백업이 있을 때 `create_backup(report.hwp, max_files=3)` 후 다른 파일의 백업 5개가 모두 남아야 합니다.
  - `report.hwp` 백업 4개 중 원본 mtime을 과거로 조작한 경우에도 **생성 시각이 가장 오래된** 백업이 삭제되어야 합니다.

---

### [ISSUE-004] CLI가 `--retry`·기본 백업을 적용하지 않고, 건너뜀 집계와 빈 폴더 오류 처리가 잘못됨

- **위치:** `hwpmate/app.py` / `_run_cli_conversion` (235–328행; 261–272행 계획 인자, 305–319행 변환 루프)
- **우선순위:** Medium
- **신뢰도:** Confirmed
- **문제:**
  1. `retry_count`와 `backup_enabled`를 `build_tasks`에 넘기지만, 이 값은 `PlannedConversion`에 저장만 될 뿐 CLI 루프는 `converter.convert_file`을 **작업당 1회** 호출합니다. 재시도와 백업(`create_backup`)은 `ConversionWorker`에만 구현되어 있습니다.
  2. 루프는 `plan.tasks`에서 `status == "건너뜀"`을 찾지만, 건너뜀 작업은 `plan.skipped_tasks`에만 들어가므로 **건너뜀은 항상 0건**으로 출력됩니다.
  3. 입력 폴더에 변환 대상이 없으면 `build_tasks`가 `ValueError("변환할 파일이 없습니다.")`를 던지는데 CLI에서 잡지 않습니다. 커스텀 `sys.excepthook`이 로그만 남기므로, 창 없는 frozen exe에서는 사용자에게 아무 메시지 없이 비정상 종료됩니다.
- **발생 조건:** README 예시대로 CLI 사용 시 항상 발생합니다(1·2). 빈 폴더나 동일 형식만 있는 폴더를 입력하면 3이 발생합니다.
- **영향:** 배치 스크립트/작업 스케줄러 자동화에서 일시적 COM 실패가 재시도 없이 실패로 끝나고, 사용자가 기대한 원본 백업이 생성되지 않으며, 결과 요약이 부정확합니다. 빈 폴더 입력은 원인 메시지 없이 실패합니다.
- **근거:** 가짜 컨버터(항상 실패)로 `main(["--input", <폴더>, "--format", "HWPX", "--retry", "3"])`를 실행한 결과, `convert_file` 1회, `create_backup` 0회, `backup/` 미생성, 출력 "성공 0건, 실패 1건, 건너뜀 0건"(실제로는 `y.hwpx` 1건 건너뜀)이었습니다. 빈 폴더 입력은 `ValueError`가 호출자까지 전파됐습니다.
- **반증 확인:** `HWPConverter.convert_file` 내부에는 재시도/백업이 없습니다(Open→Export 단일 시도 + 내부 PDF 폴백만). `tests/test_cli_conversion.py`는 성공 경로와 인자 파싱만 검증합니다.
- **호출/영향 범위:** `main` → `_run_cli_conversion` → `TaskPlanner.build_tasks`/`resolve_output_conflicts` → `HWPConverter`. GUI 경로는 영향 없음.
- **권장 수정 방향:** CLI 루프가 `ConversionWorker`와 같은 작업 실행 로직(백업 → 런타임 경로 할당 → 재시도 → 산출물 감사 필드)을 공유하도록 QThread 비의존 함수로 추출하세요. `plan.skipped_tasks`를 집계하고, `build_tasks`의 `ValueError`는 stderr 메시지 + 종료 코드로 변환하세요. CLI 결과 JSON 출력 옵션도 고려할 수 있습니다.
- **필요한 회귀 테스트:**
  - 가짜 컨버터가 처음 2회 실패 후 성공할 때 `--retry 2`면 성공 1건과 `convert_file` 3회 호출이 기대됩니다.
  - `--no-backup` 없이 실행하면 `backup/`에 사본 1개가 생겨야 합니다.
  - `.hwpx` + `--format HWPX`면 "건너뜀 1건"이 출력되어야 합니다.
  - 빈 폴더면 종료 코드 1과 stderr 안내가 나오고 예외가 전파되지 않아야 합니다.

---

### [ISSUE-005] 업데이트 헬퍼의 부모 종료 대기가 무력화되어, 경합 시 오보고 + 이후 업데이트가 계속 실패

- **위치:** `hwpmate/app.py` / `_wait_for_parent` (190–200행), `_run_apply_update` (203–232행, 219행 상태 분류); `hwpmate/services/update_installer.py` / `apply_staged_update` (218–267행), `_validate_apply_paths` (196행); `scripts/apply_update.py` (17–27행, 55행 동일 구현)
- **우선순위:** Medium
- **신뢰도:** Confirmed (대기 무력화, 상태 오분류, 잔존 백업 차단) / Likely (실제 교체 경합 실패 빈도)
- **문제:**
  1. **대기 무력화:** Windows에서 `os.kill(pid, 0)`의 `0`은 `signal.CTRL_C_EVENT`와 같아 `GenerateConsoleCtrlEvent`로 처리됩니다. 헬퍼는 `console=False` exe를 `CREATE_NO_WINDOW`로 띄우므로 콘솔이 없어 호출이 **부모가 살아 있어도 즉시 `OSError`**를 내고, `_wait_for_parent`는 곧바로 반환합니다. 존재 확인 기능이 전혀 동작하지 않습니다. 콘솔이 있는 환경(`scripts/apply_update.py`를 터미널에서 실행)에서는 반대로 콘솔 프로세스 그룹에 Ctrl+C를 보낼 위험이 있습니다.
  2. 부모(onefile 부트로더 포함)가 exe 잠금을 풀기 전에 `os.replace(staged, target)`가 실행되면 접근 거부로 실패하고, 롤백 `os.replace(backup, target)`도 같은 이유로 실패해 `UpdateApplyError("업데이트 적용 및 롤백 복구 모두 실패…")`가 발생합니다.
  3. **상태 오분류:** 상태 판정이 `"롤백" in str(exc)`이라 "롤백 복구 모두 실패" 메시지도 `rolled_back`으로 기록되어, 다음 실행 시 "이전 버전으로 복구되었습니다" 토스트가 뜹니다.
  4. **지속 실패:** 이 경로에서는 `backup` 사본(`<exe>.v9.1.0.bak`)이 남습니다. 백업 파일명은 현재 버전으로 고정되어, 같은 버전에서 다시 업데이트하면 `_validate_apply_paths`가 `FileExistsError("백업 파일이 이미 존재합니다")`를 던져 **사용자가 `.bak`을 수동 삭제할 때까지 자동 업데이트가 계속 실패**합니다.
- **발생 조건:** 부모 앱 종료 또는 onefile 부트로더의 임시 폴더 정리가 헬퍼 부트로더 기동보다 늦는 경우(느린 디스크, 백신 검사, Qt 종료 지연). 평소에는 헬퍼 onefile 추출 지연 덕분에 우연히 성공할 가능성이 큽니다.
- **영향:** 업데이트 후 앱이 재실행되지 않고(실패 경로는 relaunch 없음), 사용자에게 잘못된 "복구됨" 안내가 표시되며, 이후 업데이트가 수동 조치 전까지 계속 실패합니다. 대상 exe는 교체 실패 시 구버전 그대로라 실행 불능 위험은 낮습니다.
- **근거:** 자체 생성한 무해한 자식 프로세스에 대해 `pythonw`(콘솔 없음) + `CREATE_NO_WINDOW` 호출자에서 `os.kill(pid, 0)`을 실행하자, 대상이 살아 있는 상태에서 0.000초 만에 `OSError(9, …, winerror 6)`이 발생했습니다. 상태 분류와 `.bak` 차단은 코드 흐름으로 확인했습니다.
- **반증 확인:** `tests/test_apply_update_script.py`는 `_wait_for_parent`를 `lambda *args: None`으로 **모킹**해 실제 동작을 검증하지 않습니다. `cleanup_update_backups`는 성공 경로에서만 호출됩니다. 부모의 `app.quit()` + `sys.exit(0)`은 exe 잠금 해제 시점을 보장하지 않습니다.
- **호출/영향 범위:** `UpdateController._apply_and_restart` → `launch_update_helper` → (헬퍼) `main` → `_run_apply_update` → `_wait_for_parent` → `apply_staged_update` → `write_update_result` → (다음 실행) `check_previous_update_result`.
- **권장 수정 방향:**
  - `OpenProcess(SYNCHRONIZE)` + `WaitForSingleObject`로 실제 종료를 기다리세요. onefile 부트로더까지 고려해 `sys.executable`을 쓰기 모드로 열 수 있을 때까지 재시도하는 방식도 있습니다.
  - 상태는 예외 타입이나 명시적 플래그로 판정하세요.
  - 롤백 실패 시 `.bak`을 정리하거나, 백업명에 고유 접미사를 붙이세요.
  - 롤백 성공 시 구버전을 재실행하세요.
  - 이미 대기가 끝난 경우 `os.replace` 자체를 짧게 재시도하세요.
- **필요한 회귀 테스트:**
  - `_wait_for_parent`를 모킹하지 않고, 살아 있는 자식 PID에 대해 콘솔 없는 프로세스에서 호출하면 타임아웃 전에 반환하지 않아야 합니다.
  - `os.replace`를 두 번 모두 `PermissionError`로 패치하면 결과 상태가 `failed`이고 `.bak`이 정리되어야 합니다.
  - 기존 `.bak`이 있을 때도 재시도가 성공 경로로 진행되어야 합니다.

---

### [ISSUE-006] 변환 중 「지금 재시작하여 적용」이 종료 보호 흐름을 우회함

- **위치:** `hwpmate/ui/main_window_controllers/update.py` / `UpdateController._apply_and_restart` (244–284행, 277–278행 `app.quit()` + `sys.exit(0)`), `_show_update_dialog` (206–218행, 비모달 `dialog.show()`), `hwpmate/ui/main_window_controllers/appearance.py` / `apply_busy_ui` (97–143행 — `update_btn` 미포함)
- **우선순위:** Medium
- **신뢰도:** Likely
- **문제:** 업데이트 버튼은 변환/계획 중에도 활성 상태이고, 업데이트 대화상자는 비모달입니다. `_apply_and_restart`는 `is_conversion_active()`를 확인하지 않고 헬퍼를 띄운 뒤 `app.quit()`과 `sys.exit(0)`을 연속 호출합니다. 변환 중 종료 확인, 워커 취소 대기, 소유 한글 프로세스 정리, 설정 저장을 담당하는 `LifecycleController.close_event`의 보호 흐름을 거친다는 보장이 없습니다. 게다가 헬퍼는 ISSUE-005 때문에 부모 종료를 기다리지도 않습니다.
- **발생 조건:** 시작 3초 후 자동 업데이트 알림 또는 🔄 버튼으로 다운로드를 시작한 뒤, 다운로드 중 변환을 시작하고, 완료 후 「지금 재시작하여 적용」을 누르는 경우.
- **영향:** 진행 중 변환 중단(결과 리포트 없음), 앱이 띄운 `Hwp.exe` 고아 프로세스 잔존, 부분 산출물 잔존, 설정 미저장 가능성.
- **근거:** `apply_busy_ui`의 비활성화 목록에 `update_btn`이 없고, `_apply_and_restart`에 상태 확인이 없음을 코드로 확인했습니다. PyQt6 슬롯 내 `SystemExit` 처리와 Qt6 `quit()`의 close 이벤트 전파는 실행으로 검증하지 않았습니다.
- **반증 확인:** `closeEvent`의 변환 중 확인 대화상자는 창 닫기 경로에만 연결되어 있고, `_apply_and_restart`는 창 닫기를 요청하지 않습니다. `UpdateDialog`에도 변환 상태 참조가 없습니다.
- **호출/영향 범위:** `UpdateDialog.apply_restart_requested` → `UpdateController._apply_and_restart` → `launch_update_helper` / `QApplication.quit` / `sys.exit`. 영향 모듈: `ConversionWorker`, `HWPConverter`(소유 PID), `LifecycleController.save_settings`.
- **권장 수정 방향:** `is_conversion_active()` 또는 `is_planning`이면 적용을 거부하고 "변환 완료 후 적용" 안내를 표시하세요(또는 적용 버튼 비활성화). 종료는 `window.close()`로 요청해 `close_event`가 수락된 경우에만 헬퍼를 실행하세요.
- **필요한 회귀 테스트:** `state.is_converting=True`에서 `_apply_and_restart` 호출 시 `launch_update_helper`와 `QApplication.quit`이 호출되지 않고 안내 토스트가 표시되어야 합니다.

---

### [ISSUE-007] 이미지/HTML 변환(같은 위치 저장)이 항상 `이름 (1).png` 형태로 저장됨

- **위치:** `hwpmate/services/artifact_policy.py` / `matches_artifact_stem` (15–27행, 구분자에 `"."` 포함), `existing_artifact_conflicts` (77–98행); `hwpmate/services/task_planner.py` / `allocate_output_path` (221행)
- **우선순위:** Medium
- **신뢰도:** Confirmed
- **문제:** 보조 산출물 충돌 검사는 출력 폴더에서 stem + 구분자(`_`, `-`, 공백, `.`, `(`)로 시작하는 모든 항목을 기존 산출물로 간주합니다. 같은 위치 저장이면 **변환 입력 원본 `report.hwp` 자체**가 `report` + `.`에 매칭되어 항상 충돌로 판정되고, 출력이 `report (1).png`로 바뀝니다. 같은 폴더의 `report_final.hwp` 같은 무관한 문서도 충돌로 잡힙니다. 두 번째 실행에서는 `report (2).png`가 새로 생깁니다.
- **발생 조건:** 기본값인 같은 위치 저장 + `PNG`/`JPG`/`BMP`/`GIF`/`HTML` 대상 + 덮어쓰기 꺼짐(기본). 사실상 이 형식을 쓰는 모든 사용자에게 발생합니다.
- **영향:** 산출물 이름이 사용자 기대(`report.png`)와 다르고, 재실행할 때마다 중복 산출물이 누적됩니다. 사전 점검에 "이름 변경 N개"가 항상 표시되어 경고의 신호 가치가 떨어집니다. 우회하려면 덮어쓰기를 켜야 하는데, 이는 ISSUE-002의 위험을 키웁니다.
- **근거:** 임시 폴더에 `report.hwp`만 두고 계획을 세우자 `PNG → report (1).png`, `HTML → report (1).html`, `PDF → report.pdf`(정상)였습니다.
- **반증 확인:** `tests/test_task_planner.py`의 PNG 충돌 테스트는 모두 입력 stem(`a.hwp`, `source.hwp`)과 출력 stem(`doc.png`)이 **다르게** 구성되어 실제 사용 조건을 검증하지 않습니다.
- **호출/영향 범위:** `allocate_output_path` → `_has_existing_output_conflict` → `existing_artifact_conflicts`. 같은 매칭 규칙을 쓰는 `iter_candidate_artifact_paths`(변환 전후 스냅샷)도 입력 원본을 후보로 포함하지만, 스냅샷 비교상 변경이 없어 성공 판정에는 영향이 없습니다.
- **권장 수정 방향:** 충돌 후보에서 `SUPPORTED_EXTENSIONS`(`.hwp`/`.hwpx`) 파일과 `backup` 폴더를 제외하고, 매칭을 대상 형식 확장자(예: `.png`)와 한글이 실제 생성하는 보조 산출물 패턴(예: `stem_###.png`, `stem_files/`)으로 좁히세요. 이미지 다중 페이지 명명 규칙은 한글 버전별 스모크로 확정하는 것이 좋습니다.
- **필요한 회귀 테스트:** 폴더에 `report.hwp`만 있을 때 PNG/HTML 계획 출력이 `report.png`/`report.html`이고 이름 변경이 0건이어야 합니다. 기존 `report_001.png`가 있을 때만 `report (1).png`로 바뀌어야 합니다.

---

## 5. Potential Functional Gaps

| 구분 | 내용 | 근거/비고 |
|------|------|----------|
| **Confirmed Gap** | 업데이트 헬퍼 exe(`update-helper-<uuid>.exe`, 앱 크기만큼)가 스테이징 폴더에 계속 누적됨 | `update_installer.py:281`에서 생성만 하고 삭제 코드 없음. 다운로드만 하고 적용하지 않은 스테이징 exe도 정리 없음 |
| **Confirmed Gap** | 업데이트 롤백·실패 후 앱을 재실행하지 않음 → 사용자에게는 앱이 그냥 사라진 것으로 보임 | `_run_apply_update`는 성공 경로에서만 `Popen(target)` |
| **Confirmed Gap** | 폴더 모드 변환 시작마다 "작업 준비 중에는 입력을 변경할 수 없습니다" 토스트가 부수적으로 표시됨 | `start_scan`이 `allow_while_planning` 판정 전에 부작용 있는 `_input_locked()`를 먼저 호출 (`file_selection/controller.py:105`) |
| **Confirmed Gap** | CLI 단일 파일 입력에 확장자 검증이 없어 `.docx` 등도 한글 COM으로 열기를 시도함 | `_run_cli_conversion`이 `file_paths=[input]`을 그대로 전달, `build_tasks` 파일 모드는 확장자 미검사 |
| **Confirmed Gap** | CLI 결과를 CSV/JSON으로 남기는 옵션 없음 (GUI는 감사 필드 저장 지원) | 자동화 용도에서 `export_method`, `created_files` 감사 추적 불가 |
| **Likely Gap** | CLI는 `SingleInstanceLock`을 사용하지 않아 GUI 변환과 동시에 실행될 수 있음 → 같은 출력 파일/한글 프로세스 경합, 소유 PID 추적 혼선 | CLAUDE.md의 단일 인스턴스 의도와 불일치. 실제 충돌 빈도는 미검증 |
| **Likely Gap** | `showEvent`마다 업데이트 확인을 예약하므로, 트레이/최소화 복원 시 "나중에"로 닫은 업데이트 대화상자가 다시 뜰 수 있음 | `main_window.py:245-256`. spontaneous show 이벤트 발생 여부는 실제 환경 미확인 |
| **Likely Gap** | 200건 재순환 중 `initialize` 실패 시 경고 로그만 남기고, 남은 모든 작업이 "초기화되지 않았습니다"로 재시도 대기(작업당 1초×재시도)를 거치며 실패함. 결과 경고에도 원인이 표시되지 않음 | `worker.py:204-220`. 재순환 실패 자체의 발생 빈도는 추정 |
| 추정 | 성공 판정 뒤 마지막 `hwp.Clear(option=1)`(`converter.py:551`)이 보호되지 않아, 여기서 COM 예외가 나면 이미 검증된 산출물이 "실패"로 집계될 수 있음 | 재시도 시 스냅샷 비교로 복구될 가능성이 커 영향은 제한적 |
| 추정 | 한글 이미지 SaveAs의 다중 페이지 명명이 `stem001.png`처럼 구분자 없이 붙는 버전이면 산출물 수집에서 누락되어 실패로 판정될 수 있음 | 구분자 목록에 숫자 없음. 한글 미설치 환경이라 미검증 |
| 추정 | 재귀 스캔이 이름이 `backup`(대소문자 무시)인 **모든** 하위 폴더를 제외하므로, 사용자가 직접 만든 `Backup` 폴더의 원본은 조용히 누락됨 | CLAUDE.md에 명시된 의도된 동작이나, 사전 점검에 제외 폴더 수 안내가 없음 |

---

## 6. Documentation Mismatches

| 문서 | 기술 내용 | 실제 구현 |
|------|----------|----------|
| README CLI 옵션 표 | `--retry`: "변환 실패 시 자동 재시도 횟수", `--no-backup`: "원본 백업 생성을 건너뜀"(=기본은 백업) | CLI는 재시도·백업 모두 수행하지 않음 (ISSUE-004) |
| README "변환 지원 포맷" NOTE | "변환 대상과 동일한 형식의 파일(예: PDF로 변환 시 **대상 폴더에 이미 존재하는 .pdf 파일**)은 자동으로 건너뜁니다" | 건너뜀은 **입력 확장자가 대상 형식과 같은 경우**(HWP/HWPX 대상)만. 기존 출력 파일은 건너뛰지 않고 `(1)` 이름 변경 또는 덮어쓰기 |
| README "변환 지원 포맷" NOTE | "암호가 걸린 문서나 손상된 파일은 사전 점검 및 변환 단계에서 안전하게 감지되어 **건너뛰고**" | 사전 점검은 존재 여부와 앞 48개 파일의 1바이트 읽기만 검사. 암호 문서는 변환 단계에서 **실패**로 집계(오류 문구에 힌트 추가) |
| README 배지 / `update_history.md` | "Tests 168 Passed", "pytest 168 passed" | 현재 **174 passed** |
| README "주요 특징" | "무결성 검증 실패 시 이전 버전으로 즉시 자동 롤백" | 해시/크기 불일치는 교체 **전에** 중단(롤백 불필요). 롤백은 교체 후 `--smoke` 실패 시. 롤백 자체 실패도 "롤백됨"으로 보고 (ISSUE-005) |
| `CLAUDE.md` §3 코드베이스 구조 | 업데이트·CLI 구성요소 언급 없음 | `services/update_manifest.py`, `services/update_installer.py`, `ui/main_window_controllers/update.py`, `ui/dialogs/update_dialog.py`, `scripts/*`, `app.py`의 CLI/`--apply-update` 모드가 존재 |
| `PROJECT_STRUCTURE_ANALYSIS.md` | 업데이트 서브시스템/CLI 모드 설명 없음 | 위와 동일 (CLAUDE.md §6 문서 동기화 체크리스트 미충족) |
| `CLAUDE.md` Spec Kit 절 | "`.specify/` 또는 `specs/`가 현재 트리에 없을 수 있음" | 둘 다 존재. 다만 `specs/001-hwp-mate-reliability-ux/tasks.md`는 없음 |
| 감사 요청의 참조 문서 | `AGENTS.md` | 저장소에 존재하지 않음 |

지원 형식 목록, 단축키(Ctrl+Enter/Esc/Ctrl+O/Ctrl+Shift+O/Delete/Ctrl+Delete/F1), `retry_count` 기본 1·최대 3, 백업 보관 1~100, `pdf_export_mode` 두 값은 README/CLAUDE.md와 구현이 일치합니다.

---

## 7. Recommended Fix Plan

### Phase 1 — Immediate

데이터 손상, 주요 기능 실패

1. **ISSUE-001** 폴더 모드 변환 시작 대기 수정: `refresh_folder_scan_for_conversion`의 2초 대기를 제거하고 `FOLDER_SCAN_WAIT_MS` 대기로 일원화. 신선한 캐시가 있으면 재스캔 생략.
2. **ISSUE-002** 출력 경로가 `.hwp`/`.hwpx` 원본 또는 계획 입력/건너뜀 경로와 같으면 `overwrite`와 무관하게 이름 변경 또는 사전 점검 차단.
3. **ISSUE-003** 백업 정리 후보를 정확한 백업 파일명 패턴으로 한정하고, 파일명 타임스탬프 기준으로 정렬.

### Phase 2 — Stability

예외 처리, 입력 검증, 재시도, 상태 관리

4. **ISSUE-005** `_wait_for_parent`를 `OpenProcess` + `WaitForSingleObject` 기반으로 교체(두 구현 모두), 교체 재시도, 상태 판정을 예외 타입 기반으로, 롤백 실패 시 `.bak` 처리, 실패/롤백 후 구버전 재실행, 헬퍼 exe 정리.
5. **ISSUE-006** 변환/계획 중 업데이트 적용 차단(`update_btn` busy 처리 또는 `_apply_and_restart` 가드), 종료는 `window.close()` 경유.
6. **ISSUE-004** CLI에 재시도·백업·런타임 경로 할당 적용, `skipped_tasks` 집계, `ValueError` → 종료 코드/메시지, 입력 확장자 검증.
7. **ISSUE-007** 보조 산출물 충돌 매칭에서 원본 문서 제외 및 대상 확장자 기반으로 축소.
8. 재순환 `initialize` 실패 시 남은 작업을 즉시 실패/중단 처리하고 결과 경고에 원인 포함.

### Phase 3 — Structural

구조 개선, 테스트 가능성, 책임 분리

9. `ConversionWorker.run`의 작업 단위 로직(백업 → 경로 할당 → 재시도 → 감사 필드)을 Qt 비의존 서비스로 추출해 GUI·CLI가 공유 (ISSUE-004 재발 방지).
10. "출력 경로 안전성"(원본 보호, 보조 산출물 매칭, 덮어쓰기 정책)을 `artifact_policy`/`task_planner` 한 곳의 명시적 규칙과 테이블 테스트로 정리.
11. `_wait_for_parent`/업데이트 적용 로직 중복(`app.py` vs `scripts/apply_update.py`) 제거.
12. CLI에도 단일 인스턴스 잠금 또는 최소한 경고 적용.
13. `CLAUDE.md` §3, `PROJECT_STRUCTURE_ANALYSIS.md`, README NOTE/CLI 표/배지, `update_history.md` 동기화.

---

## 8. Test Recommendations

### Unit

| 대상 | 입력 조건 | 기대 결과 |
|------|----------|----------|
| `TaskPlanner.allocate_output_path` (ISSUE-002) | 폴더에 `a.hwp`, `a.hwpx`, 대상 `HWP`, `overwrite=True`, 같은 위치 | 작업 출력 ≠ `a.hwp` (예: `a (1).hwp`) 또는 계획 경고에 원본 보호 차단 포함 |
| 동일 (대칭) | 대상 `HWPX`, `overwrite=True` | 출력 ≠ 원본 `a.hwpx` |
| `existing_artifact_conflicts` (ISSUE-007) | 폴더에 `report.hwp`, `report_final.hwp`만 존재, 출력 `report.png` | 충돌 목록 비어 있음, 이름 변경 0 |
| 동일 | 폴더에 `report_001.png` 존재 | 충돌 1건, 출력 `report (1).png` |
| `create_backup` (ISSUE-003) | `backup/`에 `report_1_<ts>.hwp` ×5, `create_backup(report.hwp, max_files=3)` | `report_1_*` 5개 모두 유지 |
| 동일 | `report_<ts>.hwp` ×4 (원본 mtime 역순 조작), `max_files=3` | 파일명 타임스탬프가 가장 오래된 1개만 삭제 |
| `apply_staged_update` (ISSUE-005) | `os.replace` 두 호출 모두 `PermissionError` 패치 | `UpdateApplyError`, 헬퍼 상태 `failed`(not `rolled_back`), 재시도 시 `.bak` 존재로 인한 `FileExistsError` 없음 |
| `_run_apply_update` 상태 분류 | 예외 메시지 "업데이트 적용 및 롤백 복구 모두 실패" | 결과 JSON `status == "failed"` |

### Integration

| 대상 | 입력 조건 | 기대 결과 |
|------|----------|----------|
| `_run_cli_conversion` (ISSUE-004) | 가짜 컨버터: 2회 실패 후 성공, `--retry 2` | 종료 코드 0, `convert_file` 3회 호출, 요약 "성공 1건" |
| 동일 | `--no-backup` 없음, `x.hwp` 1개 | 입력 옆 `backup/x_<ts>.hwp` 생성 |
| 동일 | `y.hwpx` + `--format HWPX` | 요약 "건너뜀 1건", 종료 코드 0 |
| 동일 | 빈 폴더 입력 | 예외 미전파, 종료 코드 1, stderr 안내 |
| `ConversionWorker` 재순환 | `CONVERTER_RECYCLE_BATCH_COUNT=2` 패치, 3번째 `initialize`에서 예외, 작업 5개 | 남은 작업 3건 즉시 실패 + 결과 경고에 재순환 실패 원인 포함 (재시도 대기 없음) |

### End-to-End (관리자 권한, 한글 설치 환경 수동/반자동)

- 하위 폴더 포함 10,000개 이상(또는 네트워크 드라이브) 폴더 선택 → 미리보기 완료 → Ctrl+Enter: **사전 점검 다이얼로그 표시**(ISSUE-001).
- `report.hwp` 1개 → PNG 변환(같은 위치, 덮어쓰기 꺼짐): 산출물 `report.png`(또는 한글 보조 산출물 규칙), `(1)` 없음. 결과 CSV `created_files`에 실제 파일 전부 기록(ISSUE-007).
- 다중 페이지 HWP → PNG/JPG: 모든 페이지 이미지가 `created_files`에 포함되고 성공 판정.
- frozen exe 업데이트: 구버전 exe 실행 → 서명된 테스트 매니페스트로 업데이트 → 「지금 재시작하여 적용」 → 새 버전 자동 재실행 + "적용 성공" 토스트, `update-helper-*.exe` 잔존 없음(ISSUE-005).

### Concurrency

- 업데이트 다운로드 완료 상태에서 변환 시작 → 「지금 재시작하여 적용」 클릭: 적용 거부 안내, 변환 계속, 앱 소유 `Hwp.exe` 고아 없음(ISSUE-006).
- 헬퍼 경합: 부모 PID 프로세스가 5초간 exe 파일 핸들을 유지하도록 모의 → 헬퍼가 5초 이상 대기 후 교체 성공(ISSUE-005).
- GUI 변환 진행 중 같은 폴더 대상으로 CLI 실행: 잠금 또는 명확한 거부 메시지(Likely Gap).
- 폴더 스캔 진행 중 Ctrl+Enter 연타: 워커 1개만 시작, `is_planning` 해제 누락 없음(기존 가드 회귀).

### Regression

- 느린 스캔 재현: `iter_supported_files`를 3초 지연 패치 + 폴더 모드 `start_conversion` 1회 → `_confirm_preflight_and_start_worker` 호출, 경고 박스 0회, "작업 준비 중에는 입력을 변경할 수 없습니다" 토스트 0회(ISSUE-001 + Confirmed Gap).
- 기존 보장 유지: SaveAs 2→3 인자 폴백, `%PDF` 매직 실패 시 PrintToPDF 1회 폴백, `export_method` 기록, 소유 PID만 강제 종료, 동일 형식만 선택 시 건너뜀 전용 결과 다이얼로그.

### Platform-specific (Windows)

- `_wait_for_parent`를 콘솔 없는 프로세스(`pythonw` + `CREATE_NO_WINDOW`)에서 살아 있는 PID 대상으로 호출 → 부모 종료 전 반환하지 않음. 콘솔 있는 프로세스에서 호출해도 대상/자신에게 Ctrl+C가 전달되지 않음.
- 260자 이상 경로: 사전 점검 차단. 240~259자: `\\?\` 확장 경로 Open/SaveAs 성공.
- 대소문자만 다른 백업 파일(`Report_<ts>.hwp` vs `report.hwp`)이 정리 대상에 올바르게 포함/제외됨.
- UNC 경로(`\\server\share\...`) 입력: `to_extended_win_path`가 `\\?\UNC\...`로 변환되고 백업/출력 폴더 생성 성공.

---

## 9. Final Assessment

| 항목 | 평가 | 근거 |
|------|------|------|
| Functional Correctness | **Needs Work** | 대형/느린 폴더에서 폴더 모드 변환 불가(ISSUE-001), CLI 옵션 미적용(ISSUE-004), 이미지/HTML 출력명 항상 변경(ISSUE-007) |
| Runtime Stability | **Acceptable** | COM apartment 분리, 소유 PID 한정 강제 종료, 취소/재시도/재순환 구조가 견고함. 업데이트 적용 경로(ISSUE-005/006)만 경합에 취약 |
| Data Integrity | **Needs Work** | 덮어쓰기 시 원본 문서 덮어쓰기(ISSUE-002), 백업 정리의 교차 삭제(ISSUE-003). 설정/결과 원자 저장과 실패 산출물 정리는 양호 |
| Error Resilience | **Acceptable** | 대부분 경로에서 예외를 결과/경고로 전환. CLI 빈 폴더 미처리 예외, 업데이트 롤백 실패 오보고, 재순환 실패 미표시가 예외 |
| Cross-platform Robustness | **Acceptable** (Windows 전용 설계) | Windows 전용이 명시된 제품이며 긴 경로/UNC/UIPI 대응이 있음. 다만 Windows `os.kill` 의미를 POSIX처럼 가정한 코드가 존재(ISSUE-005) |
| Test Confidence | **Needs Work** | 174개 통과·pyright 0이지만, 실제 결함 조건(입력과 같은 stem, 느린 스캔, 교차 stem 백업, 콘솔 없는 헬퍼)을 다루지 않거나 핵심 함수를 모킹함(`_wait_for_parent`). 한글 COM 통합 테스트는 CI에 없음 |

### 실제로 먼저 수정할 문제 3개

1. **[ISSUE-001]** 폴더 모드 변환 시작 시 2초 스캔 대기로 대형/원격 폴더 변환이 영구 차단되는 문제 — 핵심 기능 장애
2. **[ISSUE-002]** 덮어쓰기 + HWP/HWPX 변환이 같은 stem 원본 문서를 백업 없이 덮어쓰는 문제 — 복구 불가능한 데이터 유실
3. **[ISSUE-003]** 백업 정리가 접두사를 공유하는 다른 파일의 백업을 삭제하는 문제 — 안전장치(백업) 무력화

그다음 순서로 **ISSUE-005/006**(업데이트 적용 경로)과 **ISSUE-004**(CLI 옵션 미적용)를 권장합니다.

---

## 10. 조치 결과 (2026-09-15)

### 10.1 감사 이슈 조치

| 이슈 | 조치 | 회귀 테스트 |
|------|------|------------|
| ISSUE-001 폴더 스캔 2초 초과 시 변환 불가 | `refresh_folder_scan_for_conversion`이 `FOLDER_SCAN_WAIT_MS`까지 대기, 방금 기다린 신선한 캐시 재사용, 계획 중 내부 재스캔의 불필요 토스트 제거 | `test_folder_conversion_starts_even_when_scan_is_slower_than_cancel_wait` + 3초 지연 재현 3/3 통과 |
| ISSUE-002 덮어쓰기 시 원본 HWP/HWPX 교체 | `allocate_output_path`가 `overwrite`와 무관하게 기존 `.hwp/.hwpx`를 출력으로 쓰지 않음, 사전 점검·CLI 경고 | `test_overwrite_never_targets_existing_source_hwp(x)_document`, 실제 COM E2E(원본 해시 동일) |
| ISSUE-003 백업 정리 교차 삭제 | 정확 파일명 패턴 매칭 + 파일명 타임스탬프 정렬 | `test_create_backup_prune_keeps_backups_of_other_prefixed_files`, `..._orders_by_backup_timestamp_not_source_mtime` |
| ISSUE-004 CLI 재시도·백업 미적용 등 | `task_runner.execute_task`를 GUI·CLI가 공유, 건너뜀 집계, `ValueError` 처리, 확장자 검증, 단일 인스턴스 잠금, `--report`, `--no-auto-continue` | `test_cli_*` 7건, 실제 COM CLI E2E |
| ISSUE-005 업데이트 헬퍼 대기 무력화·오보고·지속 실패 | `wait_for_process_exit`(OpenProcess/WaitForSingleObject), `wait_for_file_writable`, 교체 재시도, `UpdateApplyError.status`, 교체 실패 시 백업 정리, `unique_update_backup_path`, 롤백 후 재실행 | `test_wait_for_process_exit_really_waits_for_live_process`, `test_apply_staged_update_*`, `test_apply_update_script_*`, `test_app_apply_update_relaunches_after_rollback` |
| ISSUE-006 변환 중 업데이트 적용 | 변환/계획 중 적용 보류(다운로드 파일 유지), `window.close()` 흐름 경유, `sys.exit` 제거 | `test_update_apply_is_refused_while_converting` |
| ISSUE-007 이미지/HTML 이름 항상 `(1)` | 원본 문서·`backup/` 제외, 이미지 매칭을 실제 한글 명명으로 축소 | `test_image_and_html_same_location_keep_plain_output_name` 등 |

§5 Gap 조치: 업데이트 헬퍼·스테이징 잔여 파일 정리, 롤백 후 재실행, 계획 중 토스트, CLI 확장자 검증·결과 저장·단일 인스턴스, `showEvent` 반복 업데이트 확인(1회로 제한), 재순환 초기화 실패(3회 재시도 후 남은 작업 즉시 실패 + 결과 경고), 성공 후 `Clear` 예외 보호, 이미지 다중 페이지 명명(실측 확정 후 수정). 사용자 `Backup` 폴더 재귀 제외는 CLAUDE.md의 의도된 동작이므로 변경하지 않았습니다.

### 10.2 실제 한글 COM 점검에서 추가 발견·수정한 결함

환경: 한컴오피스 한글 2022 **12.0.0.4605**, ProgID `HWPFrame.HwpObject`, 비관리자 세션(`--allow-non-admin`).

| 이슈 | 우선순위 | 발견 내용 (실측) | 조치 |
|------|---------|----------------|------|
| ISSUE-008 이미지 변환 전부 실패 | High | 한글은 `sample.png` 요청 시 `sample001.png`, `sample002.png`만 생성한다. 앱은 구분자 있는 이름만 산출물로 인정해 PNG/JPG/BMP/GIF 4형식 모두 "출력 파일이 생성되지 않았습니다"로 실패 처리했고, 생성 파일도 정리하지 않고 남겼다 | `{stem}NNN`/`{stem}_NNN`(3자리) 인정, 형제 문서 산출물 오인 방지 |
| ISSUE-009 ODT 변환 실패 | Medium | `SaveAs(path, "ODT", "")` → False. `"ODF"`로 저장하면 유효한 `.odt`(ZIP) 생성 | `FormatSpec.save_format_candidates` = `ODF` → `ODT`, 형식별 2→3 인자 폴백 유지 |
| ISSUE-010 `cleanup()`의 `Clear(3)` | High (잠재 데이터 변경) | 원본 복사본으로 실험한 결과, 편집된 문서에서 `Clear(2)`/`Clear(3)` 모두 **원본 파일이 디스크에 저장됨**. `SaveAs(PDF)` 후에도 문서 경로는 원본 `.hwp`로 유지 | `cleanup()`을 `Clear(1)`(버림)로 변경 |
| ISSUE-011 DOCX/RTF 변환 무한 대기 | High | 표·각주 등이 있는 문서의 DOCX/RTF 저장 시 WPF `MessageBoxImpl` 「호환 문서 — 배치가 변경될 수 있습니다. 저장을 계속할까요? [계속 ALT+Y] [취소 ALT+N]」가 뜬다. `SetMessageBoxMode`(0x1, 0x10001, 0x1001, 0) 모두 억제 실패로 SaveAs가 무한 대기(15~20초 후 강제 종료로 확인). 창 핸들에 Y 키 전달 → 0.2초 내 저장 성공(2/2), N 키 → 저장 취소 | `HwpDialogAutoResponder`: 소유 PID + 제목 정확 일치 창에만 Y 전송, 쿨다운, 결과 경고에 횟수 표시, 설정 `auto_continue_compat_dialog`/CLI `--no-auto-continue` |
| ISSUE-012 재연결 실패 시 고아 한글 프로세스 | Medium | 강제 종료 직후 `Dispatch`는 성공하지만 첫 호출에서 "알 수 없는 인터페이스" 오류가 나고, 새 `Hwp.exe`가 남는다(점검 중 2개 발생, 시작 시각·DCOM 부모로 확인 후 정리) | `initialize` 실패 시 이번 시도에서 새로 생긴 한글 PID만 종료, 재순환 초기화 재시도 |

부가 관찰: 이 버전에서 `SaveAs` 2-인자 호출은 항상 "매개 변수 개수가 잘못되었습니다" 예외를 내고, 모든 성공이 `saveas_3`이었습니다(2→3 폴백 유지 필요성 확인). HTML 본문 이미지는 `PIC388B.png`처럼 문서 이름과 무관하게 생성되어, 변경 감지용 산출물로 기록하도록 했습니다.

### 10.3 검증

| 항목 | 결과 |
|------|------|
| `python -m pytest -q` | **222 passed** (감사 시 174 → 48개 추가) |
| `pyright .` | **0 errors** |
| 수정 전 COM 스모크 (2쪽 문서, 11개 형식) | 6/11 성공 (PNG·JPG·BMP·GIF·ODT 실패) |
| 수정 후 COM 스모크 (2쪽 문서, 11개 형식) | **11/11 성공** |
| 수정 후 COM 스모크 (복합 문서: 표·각주·다단·수식 → DOCX·RTF·ODT·PDF·PNG) | **5/5 성공**, 호환 문서 창 자동 계속 로그 확인 (사람 클릭 없음) |
| GUI 워커 E2E (실제 COM, 4개 입력 × HWP[덮어쓰기]·PNG·DOCX·ODT) | **PASS**: 원본 보호 1건, 이미지 이름 정상, 자동 계속 5회, 원본 해시 전부 동일, 백업 16개 |
| CLI E2E (실제 COM, RTF, `--output`/`--retry`/`--report`) | **성공 3/3**, 백업 3개, 자동 계속 3회, JSON 결과 저장, 종료 코드 0 |

### 10.4 남은 확인 사항

- 관리자 권한 GUI 실행, 네이티브 드래그 앤 드롭, 실제 트레이/최소화 복원 이벤트는 수동 확인이 필요합니다 (`HWP_COM_SMOKE_TEST_CHECKLIST.md` "한글 2022 실측 결함 회귀").
- PyInstaller onefile 빌드와 frozen exe 업데이트 교체 E2E는 실행하지 않았습니다. 새 모듈은 `hwp_converter.spec` hiddenimports에 추가했습니다.
- 한글 2018/2020/2024의 이미지 명명·ODT 형식 문자열·호환 문서 창 제목은 미검증입니다. 다른 제목의 확인 창은 자동 응답하지 않으므로, 발견 시 `COMPAT_DIALOG_TITLES`에 실측 후 추가해야 합니다.
- 1000쪽 이상 문서의 이미지(4자리 페이지 번호로 추정)는 산출물 목록에 누락될 수 있습니다(변환 성공 판정에는 영향 없음).

