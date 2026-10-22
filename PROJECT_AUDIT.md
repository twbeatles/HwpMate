# Project Audit

- 감사 일자: 2026-10-22
- 대상 커밋: `ea7348b` (main, 앱 버전 `9.2.3`, Qt 바인딩 PySide6 6.12.0)
- 감사 범위: `hwptopdf-hwpx_v4.py` → `hwpmate/` 패키지 전체(신규 `hwpmate/mcp/` 포함), `scripts/`, `tools/`, `tests/`, 빌드·릴리즈 설정. `legacy/`는 제외.
- 원칙: 코드 수정 없음. 반증되지 않은 문제만 이슈로 기록. 재현 스크립트는 저장소 밖 임시 폴더(세션 scratchpad)에서만 실행.
> **조치 현황 (2026-10-22 갱신):** ISSUE-001~005와 §5·§6의 항목을 모두 반영했습니다. 상세 결과와 검증은 **§10 조치 결과**를 참고하세요. 아래 §1~§9는 감사 당시 기록을 그대로 유지합니다.

- 이전 감사: 2026-09-15 감사(ISSUE-001~012)와 조치 기록은 git 이력(`106e1ef` 시점의 `PROJECT_AUDIT.md`)에 있습니다. 이번 감사에서 해당 수정이 현재 코드에 유지되는지 다시 확인했고, 회귀는 발견하지 못했습니다(§3 참고).

---

## 1. Executive Summary

### 전체 상태

GUI·CLI 변환의 핵심 경로는 이전 감사 이후 상태가 좋습니다. 원본 문서 보호(`allocate_output_path`), 정확 패턴 백업 정리, `Clear(1)` 고정, 소유 PID 한정 강제 종료, SaveAs 2→3 인자 폴백, `%PDF` 매직 검증, 업데이트 헬퍼의 실제 부모 종료 대기·롤백이 모두 코드와 테스트로 유지되고 있습니다. `pytest` 326개가 통과했고 `pyright`는 0 errors입니다.

이번 감사에서 확인한 문제는 대부분 **v9.2.0에서 추가된 MCP 서버(`hwpmate/mcp/`)** 에 있습니다. MCP 테스트는 가짜 `.py` CLI(`tests/mcp/fake_cli.py`)만 사용하므로, 배포 exe의 UAC 매니페스트나 실제 한글 프로세스 같은 현장 조건을 다루지 못합니다.

### 전체 위험도: **Medium**

GUI/CLI 사용자에게는 Low, MCP 사용자에게는 Medium~High입니다.

### 가장 중요한 문제

| # | 문제 | 우선순위 | 신뢰도 |
|---|------|---------|--------|
| ISSUE-001 | 문서 예시대로 배포 exe를 `cli_executable`로 지정하면, 관리자 권한이 아닌 MCP 호스트에서는 **모든 변환 job이 `WinError 740`으로 실패** | High | Confirmed (재현) |
| ISSUE-002 | `input_roots`/`output_roots`를 TOML에서 리스트가 아닌 **문자열로 쓰면 허용 범위가 드라이브 전체로 열림** (fail-open) | Medium | Confirmed (재현) |
| ISSUE-003 | CLI 제한 시간 초과 시 CLI만 종료되고 **한글(Hwp.exe) 프로세스가 고아로 남음**. job은 다음 파일로 계속 진행 | Medium | Likely |
| ISSUE-004 | job 도중 GUI가 실행되면(BUSY) **남은 입력이 결과에서 사라지고** 상태가 일반 실패 코드로 보고됨 | Low | Confirmed (코드 흐름) |
| ISSUE-005 | `config.toml` 문법 오류 시 MCP 서버가 **처리되지 않은 예외로 시작 단계에서 종료** | Low | Confirmed (재현) |

### 데이터 손상/유실 가능성

**원본 문서의 직접 손상·유실 경로는 발견하지 못했습니다.**
- 변환은 `Clear(1)`(버림)만 쓰고, 출력은 기존 `.hwp/.hwpx`를 덮어쓰지 않으며, MCP는 항상 `overwrite=false`로 CLI를 호출합니다.
- ISSUE-002는 데이터 파괴가 아니라 **접근 범위 위반**입니다. 허용 루트 밖의 문서를 읽어 변환하고, 허용 루트 밖에 새 파일과 `backup/` 사본을 만들 수 있습니다. 기존 파일을 덮어쓰지는 않습니다.

### 가장 먼저 수정해야 할 영역

1. MCP CLI 실행 경계: 권한 상승 필요(740) 감지와 실행 파일 형태 정책 (`mcp/runner.py`, 문서)
2. MCP 설정 로드의 타입 검증 (`mcp/config.py: McpConfig.from_mapping` / `load_config`)
3. MCP 타임아웃 후 한글 프로세스 정리 (`mcp/jobs.py: _execute`, CLI 종료 경로)

---

## 2. Project Understanding

### 프로젝트 목적

한컴오피스 한글이 설치된 Windows 10/11에서 HWP/HWPX를 PDF·DOCX·HWP/HWPX·ODT·HTML·RTF·TXT·PNG/JPG/BMP/GIF로 일괄 변환하는 PySide6 GUI + 헤드리스 CLI + 로컬 MCP(stdio) 서버입니다. 한글 COM 자동화를 사용하며 GUI는 관리자 권한 실행을 전제합니다.

### 주요 Entrypoint

| Entrypoint | 위치 | 역할 |
|-----------|------|------|
| GUI | `hwptopdf-hwpx_v4.py` → `hwpmate/app.py:main` | pywin32·관리자 확인 → `SingleInstanceLock` → `MainWindow` |
| CLI 변환 | `hwpmate/app.py:_run_cli_conversion` (`--input`) | 계획 → `execute_task` 루프 → `--report` |
| 스모크 | `hwpmate/app.py:_run_smoke` (`--smoke`) | import 무결성 JSON |
| 업데이트 헬퍼 | `hwpmate/app.py:_run_apply_update` (`--apply-update`) | 부모 종료 대기 → 이름 변경 2회 교체 → 스모크 → 롤백 |
| MCP 서버 | `python -m hwpmate.mcp` → `mcp/server.py:create_server` | preview → submit → status/result → validate |
| 보조 스크립트 | `scripts/*.py`, `tools/hwp_com_smoke.py` | 릴리즈·수동 검증 |

### 핵심 모듈

- `services/task_planner.py`: 작업 계획, 동일 형식 건너뜀, 출력 충돌 회피, 원본 보호
- `services/hwp_converter/converter.py`: COM 초기화, 보안 모듈 등록, `convert_file`, 소유 PID 종료
- `services/hwp_print_settings/`: PrintMethod 리셋, PrintToPDFEx/RunToPDF
- `workers/conversion_worker/task_runner.py`: GUI·CLI 공유 실행기 (백업 → 경로 할당 → 재시도 → 산출물 재확인)
- `services/update_manifest.py`, `services/update_installer.py`: Ed25519 서명 검증, 스트리밍 다운로드, 교체·롤백
- `mcp/planner_adapter.py`(읽기 전용 미리보기), `mcp/policy.py`(경로 허용 정책), `mcp/jobs.py`(단일 worker 큐), `mcp/runner.py`(CLI subprocess), `mcp/validator.py`(산출물 서명 검증)

### 데이터 저장 방식 (DB 없음, 모두 파일)

| 데이터 | 위치 | 쓰기 방식 |
|--------|------|----------|
| GUI 설정 | `~/.hwp_converter_config.json` | 임시 파일 → `replace`, 손상 시 `.bak` 이동 |
| 원본 백업 | 입력 파일 옆 `backup/` | `copy2` + 정확 패턴 보관 개수 정리 |
| 결과 CSV/JSON/TXT, CLI `--report` | 사용자 지정 경로 | 임시 파일 → `replace` |
| MCP 설정 | `%LOCALAPPDATA%\HwpMate\mcp\config.toml` + `HWPMATE_MCP_*` 환경 변수 | 읽기 전용 |
| MCP job 기록 | `%LOCALAPPDATA%\HwpMate\mcp\jobs\{job_id}.json`, `idempotency.json` | job 기록은 `.tmp` → `replace`, idempotency는 직접 덮어쓰기 |
| 업데이트 스테이징·결과 | `%LOCALAPPDATA%\HwpMate\updates` | `open("xb")`, 결과 JSON 원자 저장 |
| 단일 인스턴스 잠금 | `%LOCALAPPDATA%\HwpMate\HwpMate.lock` | `QLockFile` |

### 외부 의존성

한컴오피스 한글 COM(필수), pywin32, PySide6 6.12.0, cryptography(Ed25519), `mcp==2.3.0`(MCP 전용), GitHub raw/Releases(HTTPS). 빌드는 PyInstaller onefile(`console=False`, `uac_admin=True`). CI는 Python 3.15, 릴리즈 빌드는 Python 3.12입니다.

### 핵심 실행 흐름

**GUI 변환**

`변환 시작` → `ConversionController.start_conversion` → (`is_planning` 잠금, 스캔 완료 대기 `FOLDER_SCAN_WAIT_MS`) → `TaskPlanner.build_tasks` → `resolve_output_conflicts` → `PreflightDialog` → `ConversionWorker.run`(CoInitialize) → `HWPConverter.initialize(manage_com_apartment=False)` → 작업별 `execute_task` → 200건마다 재순환 → `ConversionSummary` → `ResultDialog` → CSV/JSON 원자 저장

**CLI 변환**

`--input` → `_run_cli_conversion` → `build_tasks` → `SingleInstanceLock.try_lock` → `HWPConverter.initialize()` → 작업별 `execute_task`(백업·재시도 포함) → `converter.cleanup()` → 잠금 해제 → `--report` 원자 저장 → exit code

**MCP 변환**

`hwpmate_preview_conversion` → `check_output_dir`/`check_input_path`(루트 포함 검사, reparse·UNC 거부) → `_collect_inputs` → `TaskPlanner.resolve_output_conflicts` → `plan_id`(메모리, TTL) → `hwpmate_submit_conversion`(승인 토큰, idempotency, `get_plan`의 stat 재검증) → `JobManager.submit` → worker 스레드 `_execute` → **입력 파일마다** `run_cli`(`subprocess.run`, `shell=False`, `--retry 1 --pdf-export-mode saveas_first`) → `read_cli_report` → job 기록 갱신 → `hwpmate_validate_artifacts`(존재·크기·확장자·매직)

**자동 업데이트**

`UpdateCheckWorker`(서명 검증) → `UpdateDownloadWorker`(SHA-256·크기) → `_apply_and_restart`(변환/계획 중 보류) → 헬퍼 실행 → `window.close()` → 헬퍼 `wait_for_process_exit` → `apply_staged_update`(설치본→백업, 새 파일→설치본, `--smoke`, 실패 시 롤백) → 결과 JSON → 다음 실행 시 토스트

---

## 3. Audit Coverage & Limitations

### 실제 확인한 주요 모듈 (소스 직접 열람)

- `hwpmate/mcp/` 전체 (`jobs.py`, `runner.py`, `policy.py`, `config.py`, `planner_adapter.py`, `server.py`, `__main__.py`, `report_reader.py`, `validator.py`, `capabilities.py`)
- `hwpmate/app.py`(CLI·업데이트 헬퍼), `workers/conversion_worker/task_runner.py`, `services/hwp_converter/converter.py`
- `services/update_installer.py`, `ui/main_window_controllers/update.py`, `config_repository.py`, `app_instance.py`, `windows_integration/native_drop.py`(일부)
- 이전 감사 이후 변경분 diff (`bf3fe22..HEAD`): conversion/file_selection 컨트롤러, worker, scan worker, models, config

### CodeGraph로 분석한 호출 관계

CodeGraph MCP 서버는 이번 세션에서 연결에 실패했습니다. 대신 셸 CLI(`codegraph explore`, 인덱스 139 파일 / 2,415 노드)를 사용했습니다.
- `JobManager`(24 callers, `mcp/server.py`·`mcp/validator.py`) → `_start_locked` → `_run_job` → `_execute` → `resolve_cli_executable` / `build_cli_argv` / `run_cli` → `read_cli_report` → `classify_job`
- `validate_job_outputs` ← `server.hwpmate_validate_artifacts`
- `UpdateDownloadWorker.run` → `stream_update_artifact` → `prepare_staged_update`
- CLI 쪽 caller/callee(`_run_cli_conversion` → `execute_task` → `HWPConverter.convert_file`)는 CodeGraph 결과가 부족해 파일 직접 열람으로 보완했습니다.

### 실행한 테스트/검증

| 항목 | 결과 |
|------|------|
| `python -m pytest -q` (Python 3.15.0) | **326 passed** (113s) |
| `pyright .` | **0 errors, 0 warnings** |
| 재현: `McpConfig.from_mapping({"input_roots": "<경로 문자열>"})` 및 동일 내용 TOML 로드 후 `check_input_path`/`check_output_dir` | 허용 루트 밖 파일과 `C:\Windows\win.ini`를 **허용** (ISSUE-002) |
| 재현: 문법 오류 TOML로 `load_config` | `TOMLDecodeError` 전파 (ISSUE-005) |
| 재현: 비관리자 셸에서 `subprocess.run(["dist\\HWP변환기_v9.2.3.exe", "--smoke"])` (MCP `run_cli`와 같은 호출 방식) | `OSError [WinError 740] 권한 상승이 필요합니다` (ISSUE-001) |

### 확인하지 못한 환경/외부 서비스

- 실제 한글 COM 변환은 이번 감사에서 실행하지 않았습니다. ISSUE-003의 고아 프로세스는 Windows COM 활성화 방식에 근거한 추론입니다.
- 관리자 권한 GUI 실행, 드래그 앤 드롭, 실제 업데이트 교체(서명된 릴리즈)는 수동 확인하지 않았습니다.
- GitHub 매니페스트/릴리즈 네트워크 경로는 호출하지 않았습니다.

### 반증되어 이슈에서 제외한 후보

| 후보 | 제외 이유 |
|------|----------|
| `QLockFile.setStaleLockTime(30_000)`: 30초 넘게 실행 중인 GUI의 잠금을 두 번째 인스턴스가 stale로 보고 빼앗을 수 있음 | Windows의 Qt 구현은 잠금 파일 핸들을 공유 삭제 없이 열어 두므로 stale 판정 후 삭제가 실패합니다. 잠금을 가져갈 수 없습니다. |
| MCP 파일별 CLI 호출 시 계획 출력명과 실제 출력명 불일치 | 동명 파일 2건은 CLI가 순서대로 `a.pdf`, `a (1).pdf`로 할당해 미리보기와 같은 결과가 나옵니다. MCP는 항상 `overwrite=false`입니다. |
| `json.dumps(self._idem)`이 다른 스레드의 dict 삽입과 겹쳐 `RuntimeError` 발생 | CPython C 인코더는 GIL 아래에서 items를 먼저 복사합니다. 현실적으로 발생하지 않습니다. 파일 쓰기 경합은 §5에 추정으로만 기록했습니다. |
| 업데이트 헬퍼가 `os.getpid()`(onefile 자식)만 기다려 부트로더가 exe를 잡고 있는 동안 교체 실패 | 교체는 이름 변경 방식이고 실행 중인 이미지도 이름 변경은 허용됩니다. `wait_for_file_writable`과 20회 재시도도 있습니다. |
| 업데이트 스모크 실행이 UAC 740으로 실패 | 헬퍼는 관리자 권한 GUI가 실행하므로 권한 상승 상태를 물려받습니다(GUI는 관리자 아니면 시작하지 않음). |
| PySide6 전환 잔재(PyQt6 전용 API) | `PyQt6`/`pyqtSignal`/`sip` import는 없습니다. `native_drop.py`의 `sip` 언급은 주석뿐이며 eventType 정규화 코드가 있습니다. |
| 이전 감사 ISSUE-001~012 회귀 | 스캔 대기(`FOLDER_SCAN_WAIT_MS`), 원본 보호, 정확 패턴 백업 정리, CLI `--retry`·백업, `wait_for_process_exit`, 변환 중 업데이트 보류 모두 현재 코드에 유지되고 P0 회귀 테스트가 통과합니다. |

### 분석상의 한계

- GUI 컨트롤러는 이번에 변경분 diff 위주로 확인했습니다. 변경되지 않은 부분은 이전 감사 결과와 통과한 테스트에 의존합니다.
- CodeGraph는 MCP 연결 실패로 CLI로만 사용했고, 동적 Qt 시그널 연결은 그래프에 충분히 나타나지 않았습니다.

---

## 4. High-Risk Issues

### [ISSUE-001] 문서대로 배포 exe를 `cli_executable`로 쓰면 비관리자 MCP 호스트에서 모든 변환이 실패함

- **위치:** `hwpmate/mcp/runner.py` / `run_cli`, `resolve_cli_executable`; `hwp_converter.spec`(`uac_admin=True`); `docs/MCP.md` 설정 예시
- **우선순위:** High
- **신뢰도:** Confirmed (호출 방식 재현)
- **문제:** 배포 exe는 `requireAdministrator` 매니페스트(`uac_admin=True`)로 빌드됩니다. `subprocess.run`(CreateProcess)은 UAC 승격 창을 띄우지 않고 `ERROR_ELEVATION_REQUIRED(740)`로 실패합니다. `run_cli`는 이를 `McpError("INTERNAL", "CLI 실행 실패: ...")`로 바꾸고, `_execute`에는 이를 잡는 코드가 없어 job 전체가 첫 파일에서 `failed`(INTERNAL)로 끝납니다.
- **발생 조건:** `docs/MCP.md` 예시처럼 `cli_executable = "...HwpMate-*.exe"`로 설정하고, MCP 호스트(Claude Code, Codex 등)가 일반 권한으로 실행될 때. 일반적인 사용 환경입니다.
- **영향:** MCP 변환 기능 전체가 동작하지 않습니다. 오류 코드가 `INTERNAL`이라 원인(권한)을 알기 어렵습니다. `hwpmate_get_capabilities`는 파일 존재만 보고 `cli_path_valid: true`를 반환해 실행 가능하다고 잘못 안내합니다.
- **근거:** 비관리자 셸에서 `subprocess.run([r"dist\HWP변환기_v9.2.3.exe", "--smoke"], shell=False)` → `OSError [WinError 740]`. `runner.py:116-117`에서 `OSError` → `McpError("INTERNAL")`, `jobs.py:_run_job`의 `except McpError`에서 `state="failed"`가 됩니다.
- **반증 확인:** (1) MCP 테스트는 `.py` 가짜 CLI를 `sys.executable`로 실행하므로 매니페스트 경로를 다루지 않습니다. (2) `docs/quality/HWPMATE_TEST_VERIFICATION_REPORT.md`의 MCP 실측은 "상용 CLI 래퍼"로 기록되어 있고 exe 직접 호출 결과는 없습니다. (3) MCP 호스트가 관리자 권한이면 동작하지만, 문서에 그런 요구가 없습니다.
- **호출/영향 범위:** `server.hwpmate_submit_conversion` → `JobManager.submit` → `_run_job` → `_execute` → `run_cli`. 영향은 MCP 변환 경로 전체이고, GUI/CLI 직접 사용에는 영향이 없습니다.
- **권장 수정 방향:** `run_cli`에서 `winerror == 740`을 `ELEVATION_REQUIRED` 같은 전용 코드로 분리해 보고합니다. `gather_capabilities`에서 매니페스트 또는 실행 가능 여부를 점검합니다. 배포 측에서는 `asInvoker` 콘솔 CLI 빌드를 별도로 제공하거나, 문서를 "소스 `.py` 진입점 또는 관리자 권한 MCP 호스트 필요"로 고칩니다.
- **필요한 회귀 테스트:** `subprocess.run`이 `OSError(winerror=740)`을 내도록 monkeypatch한 뒤 job이 `failed` + `ELEVATION_REQUIRED`가 되는지 확인. capabilities가 승격 필요 exe를 `cli_path_valid=false` 또는 경고로 보고하는지 확인.

### [ISSUE-002] `input_roots`/`output_roots`를 문자열로 설정하면 허용 범위가 드라이브 전체로 열림

- **위치:** `hwpmate/mcp/config.py` / `McpConfig.from_mapping`, `load_config`; `hwpmate/mcp/policy.py` / `check_containment`
- **우선순위:** Medium
- **신뢰도:** Confirmed (재현)
- **문제:** `from_mapping`은 타입 검증을 하지 않습니다. TOML에서 `input_roots = "C:/Work/HWP-Input"`처럼 문자열로 쓰면 `check_containment`의 `for root in roots`가 **문자 단위로** 반복합니다. 이때 `"/"` 또는 `"\"` 문자는 `os.path.abspath`에서 현재 드라이브 루트(`C:\`)가 되어 모든 경로가 허용 루트 안으로 판정됩니다.
- **발생 조건:** 사용자가 리스트 대신 문자열로 루트를 적을 때. 바로 옆 키인 `cli_executable`이 문자열이라 실수하기 쉽습니다. 환경 변수 경로는 리스트로 변환되므로 해당되지 않습니다.
- **영향:** fail-closed가 목적인 경로 정책이 **fail-open**이 됩니다. 에이전트가 같은 드라이브의 임의 `.hwp/.hwpx`를 읽어 변환하고, 허용 범위 밖에 출력 파일과 `backup/` 사본을 만들 수 있습니다. 기존 파일 덮어쓰기는 없습니다(overwrite=false, 원본 보호).
- **근거:** 재현 스크립트 결과: `[str] ACCEPTED: ...\outside\secret.hwp`, `[str] ACCEPTED: C:\Windows\win.ini`(입력 검사), `[str] output ACCEPTED: ...\elsewhere`, TOML 파일로도 동일하게 `[toml-str] ACCEPTED`.
- **반증 확인:** reparse/UNC 검사는 루트 포함 검사와 별개라 막지 못합니다. `win.ini`는 확장자 검사에서 걸리지만 같은 드라이브의 `.hwp`는 그대로 통과합니다. 설정 로드 단계의 검증은 없습니다(`test_p1_mcp_config.py`는 리스트 입력만 다룸).
- **호출/영향 범위:** `preview_conversion` → `check_output_dir`/`_collect_inputs` → `check_input_path`; `jobs.recheck_output` → `check_output_dir`.
- **권장 수정 방향:** `from_mapping`에서 `input_roots`/`output_roots`가 `list[str]`인지 검증합니다. 문자열은 거부하거나 단일 원소 리스트로 감싸고, 잘못된 타입은 시작 시 명확한 오류로 처리합니다. 숫자·bool 필드도 함께 타입을 검증합니다. 추가로 `check_containment`에서 드라이브 루트 자체를 루트로 허용하지 않는 방어도 고려합니다.
- **필요한 회귀 테스트:** `input_roots="C:/allowed"`(문자열) 설정 → 로드 시 오류이거나, `C:/outside/a.hwp`가 `INPUT_DENIED`. `output_roots`도 동일하게 확인.

### [ISSUE-003] CLI 제한 시간 초과 후 한글 프로세스가 고아로 남고 job은 계속 진행됨

- **위치:** `hwpmate/mcp/runner.py` / `run_cli`(`subprocess.run(timeout=...)`); `hwpmate/mcp/jobs.py` / `_execute`(timeout 분기 `continue`)
- **우선순위:** Medium
- **신뢰도:** Likely (실제 한글 환경에서는 재현하지 않음)
- **문제:** 시간이 초과되면 `subprocess.run`이 CLI 프로세스만 `kill()`(TerminateProcess)합니다. 한글은 COM out-of-process 서버로 DcomLaunch 서비스가 띄우므로 CLI의 자식이 아니고, CLI의 `finally: converter.cleanup()`과 소유 PID 강제 종료도 실행되지 않습니다. 그 뒤 job은 다음 입력으로 넘어가 새 CLI를 실행합니다.
- **발생 조건:** 한글이 모달 대화상자(암호 문서, 자동 응답 대상이 아닌 확인 창 등)나 대형 문서 처리로 `job_timeout_seconds`(기본 3600초, **파일당** 적용)를 넘길 때.
- **영향:** 문서를 연 채 멈춘 Hwp.exe가 남아 메모리와 파일 핸들을 계속 점유합니다. 다음 단계 CLI에서는 그 PID가 "이전부터 있던 프로세스"로 분류되어 이후에도 강제 종료 대상이 아닙니다. 같은 원인이 반복되면 파일마다 고아가 쌓이고 파일마다 최대 1시간씩 지연됩니다(50개면 최악 50시간).
- **근거:** `runner.py:99-117`(timeout 시 stdout/stderr만 수집), `jobs.py:375-388`(`timed_out` → 실패 기록 후 `continue`). 소유 PID 정보는 CLI 프로세스 메모리에만 있습니다(`HWPConverter.owned_pids`).
- **반증 확인:** CLI의 `KeyboardInterrupt` 처리와 `finally` 정리는 TerminateProcess에서는 실행되지 않습니다. MCP 쪽에는 한글 PID를 추적하거나 정리하는 코드가 없습니다(문서상 실행 중 강제 종료를 하지 않는 정책). `converter.initialize`의 고아 정리는 "이번 시도에서 새로 생긴 PID"만 대상이라 이전 단계 고아는 처리하지 않습니다.
- **호출/영향 범위:** `JobManager._execute` → `run_cli`; 영향은 MCP job, 이후 GUI/CLI 세션의 COM 연결 안정성(남은 한글 인스턴스)입니다.
- **권장 수정 방향:** (1) CLI가 소유 PID를 report나 사이드카 파일에 기록하고, MCP가 timeout 후 해당 PID만 이미지명 재확인 후 정리합니다(기존 소유 PID 원칙 유지). (2) timeout 발생 시 남은 입력은 실행하지 않고 `timeout`으로 종료하는 정책을 검토합니다. (3) `job_timeout_seconds`가 파일당 적용된다는 점을 이름이나 문서로 명확히 합니다.
- **필요한 회귀 테스트:** sleep 모드 가짜 CLI가 자식 프로세스를 하나 띄우고 PID 파일을 남기게 한 뒤, timeout 후 해당 PID가 정리되는지 확인. 입력 3개 중 첫 번째가 timeout일 때의 정책(중단 또는 계속)을 명시적으로 검증.

### [ISSUE-004] job 중간 BUSY 발생 시 남은 입력이 결과에서 빠지고 상태가 잘못 분류됨

- **위치:** `hwpmate/mcp/jobs.py` / `_execute` (BUSY 분기 `break`, 이후 상태 분류)
- **우선순위:** Low
- **신뢰도:** Confirmed (코드 흐름)
- **문제:** MCP는 입력마다 CLI를 따로 실행하므로, job 도중 사용자가 GUI를 켜면 다음 단계 CLI가 단일 인스턴스 잠금에 걸립니다. `_execute`는 BUSY를 실패 1건으로 기록하고 `break`합니다. 남은 입력은 `items`에도 집계에도 들어가지 않습니다. 앞 단계에 성공이 있으면 `rejected_busy`가 아니라 `partially_failed` + `OUTPUT_VALIDATION_FAILED("일부 파일 변환에 실패했습니다")`로 보고됩니다.
- **발생 조건:** 여러 파일 job 실행 중 GUI 또는 다른 CLI 실행. 사용자가 진행 상황을 GUI로 보려 할 때 일어날 수 있습니다.
- **영향:** `requested`와 `success+failed+skipped+canceled` 합계가 맞지 않습니다. 에이전트는 미처리 파일 목록을 받지 못하고, 변환 실패로 오인해 잘못된 후속 조치를 할 수 있습니다. 파일 손상은 없습니다.
- **근거:** `jobs.py:391-404`(`break`), `jobs.py:430-457`(`busy_seen and success_count == 0` 조건일 때만 `rejected_busy`).
- **반증 확인:** `test_busy_external_instance`는 첫 단계부터 BUSY인 경우만 검증합니다. `to_status_dict`는 `requested=len(inputs)`를 그대로 보고하므로 불일치가 그대로 드러납니다.
- **호출/영향 범위:** `_execute` → `hwpmate_get_job_status`/`hwpmate_get_job_result`/`hwpmate_validate_artifacts`.
- **권장 수정 방향:** `break` 전에 남은 입력을 `canceled` 또는 `not_run`(사유: BUSY_EXTERNAL_INSTANCE) 항목으로 추가하고, 부분 성공 + BUSY에는 전용 오류 코드를 보고합니다.
- **필요한 회귀 테스트:** 두 번째 입력에서만 busy를 내는 가짜 CLI(입력 이름 기반 모드)로 입력 3개 job을 실행 → 항목 3개, 합계 = requested, 오류 코드가 BUSY 계열인지 확인.

### [ISSUE-005] `config.toml` 문법 오류 시 MCP 서버가 처리되지 않은 예외로 종료됨

- **위치:** `hwpmate/mcp/config.py` / `load_config`; `hwpmate/mcp/__main__.py` / `main`
- **우선순위:** Low
- **신뢰도:** Confirmed (재현)
- **문제:** `load_config`는 `OSError`만 잡습니다. `tomllib.TOMLDecodeError`(`ValueError` 하위)는 그대로 `main`까지 올라가고, `main`은 `SystemExit`만 처리하므로 traceback과 함께 프로세스가 끝납니다.
- **발생 조건:** 설정 파일 문법 오류(따옴표 누락, Windows 경로의 `\` 이스케이프 오류 등). Windows 경로를 기본 문자열(`"C:\Work"`)로 쓰면 TOML 이스케이프 오류가 나기 쉽습니다.
- **영향:** MCP 클라이언트에는 "서버 연결 실패"로만 보이고 원인은 stderr 로그에만 남습니다. fail-closed 방향이라 보안 영향은 없습니다.
- **근거:** 재현에서 `[bad toml] raised: TOMLDecodeError`.
- **반증 확인:** `test_p1_mcp_config.py`에는 정상 TOML·환경 변수 케이스만 있습니다. 상위 wrapper는 없습니다.
- **호출/영향 범위:** `__main__.main` → `load_config` → `run_stdio`.
- **권장 수정 방향:** `TOMLDecodeError`를 잡아 파일 경로와 줄 번호가 포함된 한 줄 메시지를 stderr에 쓰고 종료 코드 2로 끝냅니다. 문서에는 경로 예시를 작은따옴표(literal string)나 `/`로 쓰도록 안내합니다.
- **필요한 회귀 테스트:** 잘못된 TOML로 `main(["--config", bad])` → 예외 대신 0이 아닌 종료 코드와 명확한 메시지.

---

## 5. Potential Functional Gaps

| 항목 | 분류 | 내용 |
|------|------|------|
| MCP 설정 플래그 다수가 동작하지 않음 | Confirmed Gap | `allow_overwrite`, `allow_disable_backup`, `allow_running_cancel`, `require_preview`, `allow_remote_http`는 코드 어디에서도 참조되지 않습니다. `require_confirmation=false`도 `server.hwpmate_submit_conversion`의 무조건 검사 때문에 효과가 없습니다. 모두 안전한 쪽으로 고정되어 위험은 없지만, 사용자는 설정이 적용된다고 오해할 수 있습니다. 미지원 키는 시작 시 경고하는 편이 좋습니다. |
| 실행 시점 입력 경로 재검증 없음 | Likely Gap | 출력 폴더는 실행 직전에 `recheck_output`으로 재검증하지만, 입력은 submit 시 `get_plan`의 size/mtime 비교만 합니다. 대기 큐(최대 600초) 동안 상위 폴더가 junction으로 바뀌는 경우는 다시 검사하지 않습니다. 현실성은 낮습니다. |
| `idempotency.json` 비원자적 쓰기 | 추정 | job 기록은 `.tmp → replace`지만 `idempotency.json`은 잠금 밖에서 `write_text`로 직접 덮어씁니다. worker 스레드와 submit이 동시에 쓰면 파일이 손상될 수 있고, 다음 시작 때 `{}`로 초기화됩니다. 그러면 같은 키·같은 계획 재제출이 중복 job이 됩니다(overwrite=false라 결과는 `(1)` 사본 생성). 기록(`*.json`)에서 다시 만들 수 있으므로 복구 로직을 권장합니다. |
| MCP 승인 토큰이 고정 공개 문자열 | 추정 (설계상 한계) | `approve_non_destructive_conversion`은 에이전트가 스스로 넣을 수 있습니다. 실제 사람 승인은 MCP 호스트의 도구 승인 UI에 의존합니다. `docs/MCP_SECURITY.md`에 이 전제를 명시할 것을 권장합니다. |
| validator가 매직 확인을 위해 파일 전체를 읽음 | Confirmed Gap (효율) | `validator.py:86`의 `path.read_bytes()[:16]`은 대형 PDF·이미지를 통째로 메모리에 올립니다. `open(...).read(16)`으로 충분합니다. |
| `recover_interrupted_swap`의 실질 도달 경로 | Likely Gap | 설치본이 없어진 상태에서는 앱을 실행할 수 없어 헬퍼도 뜨지 않으므로, 이 복구는 사실상 `scripts/apply_update.py` 수동 실행에서만 동작합니다. 두 이름 변경 사이의 중단 창은 매우 짧습니다. README에 수동 복구 절차(`*.v{ver}.bak` → 원래 이름)를 적어 두면 충분합니다. |
| onefile exe CLI가 timeout으로 종료될 때 | 추정 | PyInstaller onefile은 부트로더 부모와 Python 자식 두 프로세스로 실행됩니다. `subprocess.run` timeout은 부트로더만 종료하므로 자식 CLI가 계속 실행될 수 있습니다(PyInstaller 버전별 job object 사용 여부 미확인). ISSUE-001이 해결되어 exe 경로가 실제로 쓰이게 되면 확인이 필요합니다. |

---

## 6. Documentation Mismatches

| 문서 | 내용 | 실제 구현 | 신뢰도 |
|------|------|----------|--------|
| `docs/MCP.md` "설정" | 설정 파일 `%LOCALAPPDATA%\HwpMate\HwpMate\mcp\config.toml` | `config.default_config_path()` = `%LOCALAPPDATA%\HwpMate\mcp\config.toml` (`HwpMate` 1번). 문서 경로에 두면 설정이 무시되고 루트 미설정으로 fail closed됩니다. | Confirmed |
| `docs/MCP.md` 설정 예시 | `cli_executable = ".../HwpMate-v9.2.1-arm64.exe"` | 배포 exe는 `uac_admin=True`라 비관리자 MCP 호스트에서 실행할 수 없습니다(ISSUE-001). 버전도 현재 v9.2.3과 다릅니다. | Confirmed |
| `docs/MCP.md`/`MCP_SECURITY.md` | 관리자 권한 요구 언급 없음 | exe CLI를 쓰려면 MCP 호스트가 관리자 권한이어야 합니다. | Confirmed |
| `README.md`, `CLAUDE.md`, `AGENTS.md` | "Python 3.10+" | MCP 설정 로드에 쓰는 `tomllib`은 3.11+입니다. 3.10에서는 `tomllib=None`이 되어 `config.toml`이 **조용히 무시**됩니다(fail closed). CI는 3.15, 릴리즈는 3.12로만 검증합니다. | Confirmed (3.10 동작은 코드 근거) |
| `MCP_SECURITY.md`/설정 키 | `allow_overwrite` 등 설정 키 존재 | 해당 키는 동작하지 않습니다(§5). | Confirmed |

GUI/CLI 관련 README·CLAUDE.md 설명(SaveAs 폴백, `Clear(1)`, 원본 보호, 백업 정리, CLI `--retry`/`--no-backup`/`--report`, 업데이트 대기 방식)은 현재 구현과 일치합니다.

---

## 7. Recommended Fix Plan

### Phase 1 — Immediate

1. **ISSUE-001**: `run_cli`에서 `WinError 740` → `ELEVATION_REQUIRED` 오류 코드, capabilities 경고, `docs/MCP.md` 예시와 요구 권한 정정. 장기적으로 `asInvoker` 콘솔 CLI 빌드를 검토합니다.
2. **ISSUE-002**: `McpConfig.from_mapping` 타입 검증(`input_roots`/`output_roots`는 `list[str]`만, 그 외는 시작 실패). 드라이브 루트 허용 여부도 명시적 정책으로 정합니다.

### Phase 2 — Stability

3. **ISSUE-003**: CLI가 소유 한글 PID를 기록하고 MCP가 timeout 후 해당 PID만 정리. timeout 이후 남은 입력의 처리 정책(중단 또는 계속)을 정하고 문서화합니다.
4. **ISSUE-004**: BUSY `break` 시 남은 입력을 `not_run`으로 기록하고 전용 오류 코드를 보고합니다.
5. **ISSUE-005**: TOML 파싱 오류를 명확한 종료 메시지로 처리합니다.
6. `idempotency.json` 원자 저장 + 잠금 안에서 쓰기 + 손상 시 job 기록으로 재구성.
7. `docs/MCP.md` 설정 경로 오타 수정, Python 3.11+(MCP) 요구 명시.

### Phase 3 — Structural

8. 동작하지 않는 MCP 설정 키 정리(구현하거나 제거하고, 미지원 키는 경고).
9. MCP 통합 테스트에 "매니페스트가 있는 실행 파일" 또는 OSError 주입 케이스 추가. 가짜 CLI에 입력별 모드(`busy`/`sleep`을 특정 파일에만 적용)를 추가해 다단계 시나리오를 검증합니다.
10. validator 헤더 읽기를 스트림 방식으로 변경.

---

## 8. Test Recommendations

### Unit

- `McpConfig.from_mapping({"input_roots": "C:/a"})` → `ValueError`(또는 `["C:/a"]`로 정규화). 이후 `check_input_path(cfg, "C:/b/x.hwp")` → `INPUT_DENIED`.
- `McpConfig.from_mapping({"output_roots": 5})` → 로드 실패.
- `load_config(bad_toml)` → 명확한 오류(예외 타입 또는 종료 코드), traceback 없음.
- `run_cli`에서 `subprocess.run`을 `OSError(None, "...", None, 740)`로 패치 → `McpError.code == "ELEVATION_REQUIRED"`.
- `validate_job_outputs`가 1GB 가짜 파일에서 헤더 16바이트만 읽는지(`open` 패치로 read 크기 확인).

### Integration

- 가짜 CLI 입력별 모드: 입력 `a,b,c`에서 `b`만 busy → job 항목 3개, `c`는 `not_run`, `requested == 합계`, 오류 코드 BUSY 계열.
- 입력 `a,b`에서 `a`만 sleep(timeout) → 정해진 정책대로(중단 시 `b`는 `not_run`, 계속 시 `b` 성공) 결과와 상태가 일치하는지 확인.
- `idempotency.json`을 손상시킨 뒤 `JobManager` 재생성 → 같은 키·같은 계획 제출이 기존 job을 반환하는지 확인.

### End-to-End (관리자 권한·한글 설치 환경, 수동/반자동)

- 일반 권한 Claude Code/Codex에서 `python -m hwpmate.mcp` + exe `cli_executable` → preview/submit → 수정 후 `ELEVATION_REQUIRED` 안내 확인. `.py` 진입점으로는 PDF 1건 성공 + `validate_artifacts` verified 확인.
- 암호 문서로 timeout 유도(`job_timeout_seconds=30`) → 종료 후 작업 관리자에서 Hwp.exe가 남는지 확인(수정 전 재현, 수정 후 정리 확인).
- 기존 GUI 시나리오: 폴더 드롭 → 사전 점검 → PDF/DOCX 변환 → 결과 CSV/JSON 감사 필드(`created_files`, `export_method`, `save_format`, `progid_used`) 확인.

### Concurrency

- MCP job 실행 중 GUI 실행 → GUI는 "이미 실행 중"으로 막히지 않아야 하는지(현재는 GUI가 잠금을 얻을 수 있는 단계 사이 틈이 있음) 정책을 정하고 검증합니다.
- submit 다중 스레드(10개 동시, 서로 다른 키) → `idempotency.json`이 유효한 JSON이고 모든 키를 포함하는지 확인.

### Regression

- 이전 감사 P0 테스트(`test_p0_*`) 유지: 원본 `.hwp/.hwpx` 보호, 백업 격리, CLI/GUI 동일 실행기.
- `CLAUDE.md` §2 금지 항목: `CreateAction("Print")` 미사용, `Clear(2/3)` 미사용, 전역 `Hwp.exe` 종료 미사용을 grep 기반 정적 테스트로 고정합니다.

### Platform-specific (Windows)

- `uac_admin=True` exe를 일반 권한 `subprocess`로 실행 → 740 처리 확인(CI windows-latest에서 매니페스트가 있는 더미 exe로 가능).
- Python 3.10 환경에서 MCP 설정 로드 시 `tomllib` 부재를 경고하는지 확인. 또는 지원 버전을 3.11+로 올리고 문서를 맞춥니다.
- 경로에 한글·공백·260자 초과가 있는 입력으로 MCP preview 경고와 CLI 실행 결과가 일치하는지 확인.

---

## 9. Final Assessment

| 항목 | 평가 | 근거 |
|------|------|------|
| Functional Correctness | **Acceptable** | GUI/CLI 핵심 변환 흐름은 정확합니다. MCP는 문서대로 배포 exe를 쓰면 동작하지 않습니다(ISSUE-001). |
| Runtime Stability | **Acceptable** | 단일 인스턴스, COM apartment 분리, 재순환, 업데이트 롤백이 견고합니다. MCP timeout 후 고아 한글 프로세스가 남는 문제가 있습니다(ISSUE-003). |
| Data Integrity | **Good** | 원본 덮어쓰기 방지, `Clear(1)` 고정, 원자 저장, 정확 패턴 백업 정리. 원본 손상 경로는 발견하지 못했습니다. |
| Error Resilience | **Acceptable** | GUI/CLI 오류 처리는 충분합니다. MCP는 권한·설정 오류가 `INTERNAL`이나 traceback으로 드러납니다(ISSUE-001, 005). |
| Cross-platform Robustness | **Acceptable** | Windows 전용 도구로 의도된 범위에서 경로 후보·인코딩 처리가 좋습니다. Python 3.10 지원 표기와 MCP 요구사항이 맞지 않습니다. |
| Test Confidence | **Acceptable** | 326개 테스트와 pyright 0 errors. 다만 MCP 테스트는 `.py` 가짜 CLI에만 의존해 UAC·실제 프로세스·다단계 BUSY/timeout 조합을 다루지 않습니다. |

### 실제로 먼저 수정할 문제 3개

1. **ISSUE-001**: MCP CLI 실행의 권한 상승(740) 처리와 문서·배포 형태 정정. 이 문제가 있으면 MCP 기능이 문서대로 동작하지 않습니다.
2. **ISSUE-002**: MCP 허용 루트 설정의 타입 검증. 보안 경계가 설정 실수 한 번으로 fail-open이 됩니다.
3. **ISSUE-003**: timeout 후 소유 한글 프로세스 정리와 남은 입력 처리 정책.

---

## 10. 조치 결과 (2026-10-22, v9.2.4)

### 10.1 감사 이슈 조치

| 이슈 | 조치 | 위치 | 회귀 테스트 |
|------|------|------|------------|
| ISSUE-001 권한 상승 | `Popen` 실패 시 `winerror == 740`을 `ELEVATION_REQUIRED`로 분리. capabilities가 exe 매니페스트(`requireAdministrator`)와 현재 권한을 비교해 `cli_requires_elevation`·`cli_path_valid=false` 보고. 문서 예시를 소스 진입점으로 교체하고 실행 권한 절 추가 | `mcp/runner.py`, `mcp/capabilities.py`, `docs/MCP*.md` | `test_elevation_required_*`, `test_capabilities_flag_admin_manifest` |
| ISSUE-002 허용 루트 fail-open | 문자열은 단일 원소 목록으로 감싸고, 그 외 타입·드라이브 전체 루트·필드 타입 오류는 `McpConfigError`로 시작 거부. `check_containment`도 문자열·드라이브 루트를 방어적으로 무시 | `mcp/config.py`, `mcp/policy.py` | `test_string_roots_*`, `test_drive_root_is_rejected_at_load`, `test_direct_config_with_bad_roots_still_fails_closed`, `test_wrong_types_are_rejected`(6건) |
| ISSUE-003 timeout 고아 한글 | timeout 시 `taskkill /T`로 CLI 프로세스 트리 종료. CLI 숨김 인자 `--owned-pids-file`로 소유 한글 PID 기록(초기화·재순환 후, 정상 정리 후 빈 목록), MCP가 매 단계 후 남은 소유 PID만 이미지명 재확인 후 종료. timeout 후 남은 입력은 실행하지 않음 | `mcp/runner.py`, `mcp/jobs.py`, `app.py` | `test_timeout_stops_remaining_inputs`, `test_timeout_cleans_owned_processes_outside_cli_tree`, `test_timeout_kills_process_tree`, `test_cleanup_only_targets_live_hwp_pids`, `test_normal_exit_leaves_nothing_to_clean`, `test_cli_records_and_clears_owned_pids_file` |
| ISSUE-004 BUSY 집계 | 중단 시 남은 입력을 `취소됨`(`미실행: <사유>`)으로 기록. 부분 성공 후 BUSY/timeout은 `partially_failed` + 전용 오류 코드 | `mcp/jobs.py` | `test_busy_mid_job_records_remaining_inputs` |
| ISSUE-005 TOML 오류 | `TOMLDecodeError`·형식 오류 → stderr 한 줄 + 종료 코드 2. `tomllib` 없는 Python 3.10에서 설정 파일이 있으면 시작 거부 | `mcp/config.py`, `mcp/__main__.py` | `test_bad_toml_raises_config_error_with_path`, `test_main_exits_cleanly_on_bad_config`, `test_config_file_without_tomllib_is_not_silently_ignored` |

### 10.2 §5 Gap·§6 문서 조치

| 항목 | 조치 |
|------|------|
| 적용되지 않는 MCP 설정 키 | 필드는 유지하되 `FIXED_POLICY_VALUES`와 다른 값이면 시작 로그·capabilities `warnings`에 표시. `require_confirmation=false`도 무시하고 승인 토큰을 항상 검사 |
| 실행 시점 입력 재검증 | `recheck_input`으로 각 입력을 실행 직전 허용 정책·존재 여부로 재검사, 실패 시 해당 입력만 실패 기록 (`test_queued_job_input_removed_before_execution`) |
| `idempotency.json` 비원자 쓰기 | 잠금 안에서 임시 파일 → `replace`, 시작 시 job 기록으로 항상 보강·재구성 (`test_corrupt_idempotency_index_is_rebuilt`, `test_concurrent_submits_keep_index_valid`) |
| 고정 승인 토큰 | `MCP_SECURITY.md`에 "사람 승인은 호스트의 도구 실행 승인에 의존" 전제 명시 |
| validator 전체 파일 읽기 | 16바이트만 읽도록 변경 |
| `recover_interrupted_swap` 도달 경로 | README FAQ Q8에 `*.exe.v*.bak` 수동 복구 절차 추가 |
| onefile 자식 잔존(추정) | timeout 시 `taskkill /T` 트리 종료로 함께 해소 |
| 문서 불일치 | `docs/MCP.md` 설정 경로 오타 수정, exe 예시·권한 요구, Python 3.11+(MCP) 명시(README·CLAUDE.md·AGENTS.md), 정책 키 동작 설명. `PROJECT_STRUCTURE_ANALYSIS.md`에 `hwpmate/mcp/` 구조 추가 |
| 금지 호출 정적 고정 (§8 Regression) | `tests/test_forbidden_patterns.py`: `.Execute(` 첫 인자는 `"PrintToPDFEx"`만, `.Clear(`는 1만, `taskkill`은 `/PID` 지정만 허용 |

### 10.3 검증

| 항목 | 결과 |
|------|------|
| `python -m pytest -q` (Python 3.15.0) | **360 passed** (기존 326 + 신규 34) |
| `pyright .` | 0 errors, 0 warnings |
| 실측: 일반 권한, 소스 CLI(`hwptopdf-hwpx_v4.py`), 한글 2022, 합성 문서 → MCP preview→submit→validate PDF | `succeeded`, `validate` 1/1 verified. 변환 중 PID 파일 `{"pids":[19156]}` → 종료 후 `{"pids":[]}`, 잔여 Hwp.exe 없음 |
| 실측: 586쪽 합성 문서 DOCX, `job_timeout_seconds=4` | `timeout`(TIMEOUT), 경고 "CLI가 남긴 한글 프로세스 1개를 정리했습니다 (PID 20616)", 잔여 Hwp.exe 없음 |
| 실측: capabilities에 배포 exe(`dist\HWP변환기_v9.2.3.exe`) 지정, 일반 권한 | `cli_path_valid=false`, `cli_requires_elevation=true` |

### 10.4 남은 확인 사항 / 보류

- **`asInvoker` 콘솔 CLI 별도 빌드는 보류했습니다.** 빌드 산출물·릴리즈 워크플로·업데이트 매니페스트(단일 exe 전제)가 함께 바뀌는 구조 변경이라, 이번에는 권한 감지·전용 오류 코드·문서 안내(소스 진입점 또는 관리자 호스트)로 대응했습니다. 배포 exe만으로 일반 권한 MCP를 지원하려면 별도 작업으로 진행해야 합니다.
- MCP job 실행 중 사용자가 GUI를 켜는 경우, 단계 사이 틈에서 GUI가 잠금을 얻을 수 있습니다. 이 경우 MCP는 남은 입력을 `미실행`으로 기록하고 `BUSY_EXTERNAL_INSTANCE`로 보고하는 것을 정책으로 정했습니다(GUI 차단은 하지 않음).
- 관리자 권한 GUI 수동 시나리오(드롭·사전 점검·결과 저장)는 이번 변경 범위 밖이라 재확인하지 않았습니다.
