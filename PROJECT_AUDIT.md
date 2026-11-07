# Project Audit

- **대상:** HwpMate v9.2.4 (`main` @ `4a5cc57`)
- **감사일:** 2026-11-07
- **범위:** 기능 구현·런타임 안정성 (코드 수정 없음)
- **조치:** 2026-11-07 같은 날 v9.2.5에서 모든 이슈·기능 공백·문서 불일치를 수정했습니다. 결과는 [§10 수정 결과](#10-수정-결과-2026-11-07)에 있습니다. §1~§9는 수정 전 감사 당시 기준입니다.
- **이전 감사:** 2026-10-22 감사본(MCP 중심, ISSUE-001~005)은 v9.2.4에서 조치되었고 git 이력(`9a5cf31`)에 남아 있습니다. 이번 감사는 그때 변경분 위주로만 본 **GUI·CLI 변환 엔진, 산출물 판정, 프로세스 소유권, 업데이트·종료 흐름**을 중심으로 다시 봤습니다.

---

## 1. Executive Summary

### 전체 상태

코드베이스는 성숙하고 방어적입니다. 원본 보호(`.hwp/.hwpx` 덮어쓰기 금지), 실패 산출물 정리 시 "이전에 있던 파일은 지우지 않음" 원칙, 원자적 설정·리포트 쓰기, Ed25519 서명 업데이트와 롤백, 단일 인스턴스 잠금이 일관되게 적용되어 있습니다. 테스트 359개가 통과하고 pyright 오류는 0개입니다. 이전 감사의 MCP 이슈 5건은 모두 회귀 테스트와 함께 고쳐진 상태입니다.

### 전체 위험도: **Low ~ Medium**

데이터를 파괴하는 경로는 찾지 못했습니다. 남은 문제는 **한글 프로세스 수명 관리**에 몰려 있습니다. 어떤 프로세스를 "앱 소유"로 볼지, 그리고 한글이 죽은 뒤 어떻게 복구할지입니다.

### 가장 중요한 문제

| # | 문제 | 우선순위 | 신뢰도 |
|---|------|---------|--------|
| ISSUE-001 | 소유 PID를 "초기화 전후 PID 차집합"으로 판정해, 그 사이에 사용자가 연 한글을 앱 소유로 오인 → 사용자 창 반복 숨김·강제 종료·MCP 자동 종료 대상 | Medium | Likely |
| ISSUE-002 | 배치 중 한글 프로세스가 죽으면 다시 연결하지 않아, 다음 재순환(최대 199건 뒤)까지 남은 파일이 모두 실패 | Medium | Likely |
| ISSUE-003 | 이미지/HTML 보조 산출물 판정이 너무 넓어, 이름이 비슷한 형제 폴더·파일을 충돌로 오인해 출력 이름이 불필요하게 바뀜 | Low | Confirmed (재현) |
| ISSUE-004 | CLI "건너뜀만 있는" 경로에서 `--report` 저장이 실패하면 처리되지 않은 traceback으로 끝남 | Low | Confirmed (재현) |

### 데이터 손상/유실 가능성

- **원본 문서:** 손상·유실 경로를 찾지 못했습니다. `cleanup()`·실패 경로 모두 `Clear(1)`(저장 안 함)만 쓰고, 출력 경로가 기존 `.hwp/.hwpx`이면 덮어쓰기 설정과 무관하게 새 이름을 씁니다(`task_planner._blocks_output_path`).
- **기존 산출물:** `remove_new_attempt_artifacts`는 시도 전에 있던 파일은 지우지 않습니다. 덮어쓰기 모드에서 PDF 매직 검증에 실패한 "변경된" 파일만 지우는데, 이건 이미 한글이 덮어쓴 뒤라 추가 유실이 아닙니다.
- **사용자의 다른 한글 문서:** ISSUE-001의 조건이 맞으면 사용자가 따로 연 한글 프로세스가 강제 종료되어 **저장하지 않은 편집 내용이 사라질 수 있습니다.** GUI에서는 확인 창을 거치지만, MCP timeout 정리 경로에서는 확인 없이 자동으로 종료됩니다.

### 가장 먼저 수정해야 할 영역

`HWPConverter.initialize`의 소유 프로세스 판정(ISSUE-001)과 COM 연결 끊김 복구(ISSUE-002)입니다. 둘 다 `hwpmate/services/hwp_converter/converter.py`와 `workers/conversion_worker/task_runner.py`에 있어 함께 손보기 좋습니다.

---

## 2. Project Understanding

### 프로젝트 목적

한컴오피스 한글 COM 자동화로 HWP/HWPX를 PDF·DOCX·HWPX·HWP·HTML·ODT·TXT·RTF·이미지로 일괄 변환하는 Windows 전용 도구입니다. GUI(PySide6), 헤드리스 CLI, 로컬 stdio MCP 서버 세 가지 진입점이 같은 변환 실행기를 씁니다.

### 주요 Entrypoint

| 진입점 | 경로 | 비고 |
|--------|------|------|
| GUI | `hwptopdf-hwpx_v4.py` → `hwpmate.app.main` → `MainWindow` | 관리자 권한 필수, `SingleInstanceLock` |
| CLI | `main(--input ...)` → `_run_cli_conversion` | 관리자 검사 없음, 같은 잠금 사용 |
| 업데이트 헬퍼 | `main(--apply-update ...)` → `_run_apply_update` | 스테이징 폴더에 복사된 exe로 실행 |
| 스모크 | `main(--smoke)` → `_run_smoke` | 업데이트 후 검증에도 사용 |
| MCP | `python -m hwpmate.mcp` → `mcp/server.py` → `JobManager` → CLI 서브프로세스 | Python 3.11+, 허용 루트 정책 |

### 핵심 모듈

| 모듈 | 책임 |
|------|------|
| `services/task_planner.py` | 입력 수집, 출력 경로 계산, 충돌 회피(`allocate_output_path`), 원본 보호 |
| `services/artifact_policy.py`, `hwp_converter/artifact_snapshot.py` | 산출물 후보·충돌 판정, 변경 감지, 실패 산출물 정리 |
| `services/hwp_converter/converter.py` | COM 연결, 보안 모듈 등록, Open → (인쇄설정 리셋) → SaveAs/PrintToPDFEx → 산출물 검증, 소유 PID 추적·강제 종료 |
| `services/hwp_print_settings/` | PDF 이중 엔진, `%PDF` 매직 검사, 불완전 산출물 정리 |
| `workers/conversion_worker/task_runner.py` | 작업 1건 실행(백업 → 런타임 경로 재할당 → 변환·재시도 → 산출물 재확인), 재순환 |
| `workers/conversion_worker/worker.py` | GUI용 QThread 래퍼, 200건마다 재순환 |
| `ui/main_window_controllers/*` | 계획·사전 점검·워커 오케스트레이션, 취소/강제 종료, 종료 처리, 업데이트 |
| `windows_integration/*` | 소유 PID 창 숨김·전면화, 보안 창 자동 클릭, 「호환 문서」 자동 계속 |
| `services/update_*.py` | 매니페스트 서명 검증, 다운로드·해시 확인, 이름 변경 방식 교체, 스모크 후 롤백 |
| `mcp/*` | 정책·계획·job 큐·CLI 실행·리포트 해석·검증 |

### 데이터 저장 방식 (DB 없음)

| 데이터 | 위치 | 쓰기 방식 |
|--------|------|----------|
| GUI 설정 | `~/.hwp_converter_config.json` | 임시 파일 → `replace`, 손상 시 `.bak`으로 옮기고 기본값 사용 |
| 원본 백업 | `<입력 폴더>/backup/{stem}_{timestamp}{ext}` | `copy2`, 파일당 1~100개 보관 |
| 보안 모듈 | `%LOCALAPPDATA%\HwpMate\security\*.dll` + `HKCU\Software\HNC\Hwp{Automation,Ctrl}\Modules` | SHA-256 검증 후 설치·등록 |
| 업데이트 | `%LOCALAPPDATA%\HwpMate\updates\` | 스테이징 exe, 결과 JSON(원자적) |
| 잠금 | `%LOCALAPPDATA%\HwpMate\HwpMate.lock` | `QLockFile` |
| MCP job | `%LOCALAPPDATA%\HwpMate\mcp\jobs\` | 기록·idempotency 모두 원자적 |

### 외부 의존성

PySide6 6.12.0, pywin32, cryptography, (MCP) `mcp==2.3.0`, 한컴오피스 한글 2018+ COM(`HWPFrame.HwpObject` 등), GitHub Releases(업데이트 매니페스트·아티팩트).

### 핵심 실행 흐름

```
GUI: 변환 시작(Ctrl+Enter)
  → ConversionController.start_conversion
      → (폴더 모드) 스캔 캐시 확보/대기 → validate_output_settings
      → TaskPlanner.build_tasks → resolve_output_conflicts (계획 시점 이름 할당)
      → PreflightDialog (차단 오류 시 시작 불가) → _save_settings
  → ConversionWorker.run (QThread, CoInitialize)
      → HWPConverter.initialize (보안 모듈 준비 → Dispatch → RegisterModule → 소유 PID 산출)
      → for task: execute_task
            → create_backup → allocate_output_path(런타임 재확인) → mkdir
            → convert_file (호환 문서 자동 응답 스레드 ON)
                → Open → 인쇄설정 리셋 → SaveAs / PrintToPDFEx → 산출물 스냅샷 비교 → PDF 매직 검증
                → 실패 시 새로 생긴 산출물만 삭제, Clear(1)
            → verify_claimed_artifacts
         200건마다 recycle_converter
      → task_completed(요약) → finally cleanup(Clear(1), Quit)
  → on_task_completed → ResultDialog (실패만 재변환 / CSV·JSON 저장)

CLI: --input → TaskPlanner → SingleInstanceLock → HWPConverter → execute_task 루프 → 리포트 → 종료 코드
MCP: preview → submit(승인 토큰) → JobManager 단일 worker → 입력별 CLI 서브프로세스(--owned-pids-file) → 리포트 해석 → validate
```

---

## 3. Audit Coverage & Limitations

### 실제 확인한 주요 모듈 (소스 직접 열람)

- `services/hwp_converter/converter.py`(전체), `artifact_snapshot.py`, `process_snapshot.py`, `com_types.py`
- `services/artifact_policy.py`, `services/task_planner.py`, `services/hwp_print_settings/pdf_export.py`
- `services/hwp_security_module.py`, `services/update_installer.py`(전체), `ui/main_window_controllers/update.py`(적용 흐름)
- `workers/conversion_worker/{worker,task_runner,backup,summary,protocol}.py`, `workers/file_scan_worker.py`
- `ui/main_window_controllers/conversion/controller.py`(전체), `lifecycle/controller.py`(전체), `file_selection/controller.py`(구조), `ui/dialogs/{preflight,atomic_io}.py`
- `app.py`(전체), `app_instance.py`, `config_repository.py`, `path_utils.py`
- `windows_integration/hwp_dialog_responder.py`, `hwp_window_control.py`(구조)
- `mcp/jobs.py`(실행·복구 부분), `mcp/runner.py`(실행·PID 정리 부분)

### CodeGraph로 분석한 호출 관계

이번 세션에서는 CodeGraph MCP(`codegraph_explore`)가 정상 동작했습니다.

- `HWPConverter`(11 callers: `app.py`, `conversion_worker/worker.py`, `tools/hwp_com_smoke.py` 등) → `convert_file` → `_convert_file_impl` → `_try_saveas` / `_try_print_to_pdf` / `_artifact_ok` / `_cleanup_failed_artifacts`
- `ConverterEngine.convert_file` ← `task_runner.execute_task` ← `ConversionWorker.run`, `app._run_cli_conversion` (GUI·CLI 공유 경로 확인)
- `_terminate_hwp_pids` ← `kill_owned_processes`(GUI 강제 종료), `initialize`(실패 시 고아 정리)
- `read_owned_pids` / `cleanup_owned_hwp_processes`(`mcp/runner.py:247`) ← `JobManager._execute`(`mcp/jobs.py:420`) ← CLI `_write_owned_pids`(`app.py:416`)
- Qt 시그널 연결(`worker.task_completed` → `on_task_completed` 등)은 그래프에 충분히 나타나지 않아 파일 열람으로 보완했습니다.

### 실행한 테스트/검증

| 항목 | 결과 |
|------|------|
| `.venv` `python -m pytest -q` (Python 3.15.2) | **359 passed, 1 skipped** (110s). 기본 마커 설정상 `com_real`/`gui_real`/`release_vm`은 제외됨 |
| `pyright .` | **0 errors, 0 warnings** |
| 재현 스크립트(임시 폴더): `allocate_output_path` + `existing_artifact_conflicts` | PNG: 형제 폴더 `report_2024/` 때문에 `report.png` → `report (1).png`. 스냅샷 후보에 `report_2024/notes.hwp`(원본 문서)·`other.txt` 포함. HTML: `doc_final.html` 때문에 `doc.html` → `doc (1).html`. PDF(대조군): 이름 변경 없음 (ISSUE-003) |
| 재현: 건너뜀만 있는 폴더 + 디렉터리를 만들 수 없는 `--report` 경로로 CLI 실행 | `FileExistsError` traceback, 종료 코드 1 (ISSUE-004) |

### 확인하지 못한 환경/외부 서비스

- 실제 한글 COM 변환, 한글 프로세스 크래시, PrintToPDFEx 프린터 동작은 이번에 실행하지 않았습니다. ISSUE-001·002는 코드 흐름에 근거한 판단입니다.
- 관리자 권한 GUI 수동 조작(드롭, 강제 종료, 결과 창), 실제 서명 업데이트 교체, GitHub 네트워크 경로는 실행하지 않았습니다.
- 저장소 규칙 파일 `CLAUDE.md`/`AGENTS.md`는 `.gitignore` 대상이고 이 작업 트리에 없어서 확인하지 못했습니다. README와 `PROJECT_STRUCTURE_ANALYSIS.md`, 이전 감사본으로 대신했습니다.

### 반증되어 이슈에서 제외한 후보

| 후보 | 제외 이유 |
|------|----------|
| 덮어쓰기 모드 + PDF 검증 실패 시 기존 PDF 삭제 | 삭제는 "시도 중 바뀐" 파일만 대상이며, 한글이 이미 덮어쓴 뒤라 추가 유실이 아닙니다. 스냅샷과 같으면 지우지 않습니다(`remove_incomplete_output`). |
| 실패 정리가 형제 폴더의 원본 `.hwp`를 지울 수 있음 (ISSUE-003의 스냅샷 확대) | `remove_new_attempt_artifacts`는 시도 전에 있던 경로를 절대 지우지 않습니다. 원본은 시도 전부터 있으므로 삭제되지 않습니다. |
| 결과 창(`dialog.exec()`) 중첩 이벤트 루프 안에서 `finished`가 처리되어 상태가 꼬임 | `on_worker_finished`가 끝나기 전에는 `is_converting=True`라 "실패만 재변환"이 막히고, 끝난 뒤에는 새 워커가 정상 시작됩니다. 동시에 두 워커가 뜨는 경로는 없습니다. |
| 종료 요청 후 워커 종료 시 `close`가 두 번 호출됨 | 두 번째 `close_event`는 설정을 다시 저장할 뿐 부작용이 없습니다. |
| 자동 응답 스레드가 `owned_pids`를 복사하는 중 UI 스레드가 수정해 `RuntimeError` | `_run` 루프가 예외를 삼키고 다음 주기에 다시 시도합니다. |
| 보안 DLL 교체가 실행 중 한글에 잠겨 실패 | `OSError` 시 무결성을 확인한 기존 설치본을 씁니다. |
| 재순환 시 이전 한글이 아직 종료 중이라 새 PID 계산이 틀림 | 종료 중인 프로세스는 `before_pids`에 들어가 소유 집합에서 빠집니다(누락 방향이라 안전). |
| 이전 감사 ISSUE-001~005 및 §5 조치 회귀 | 관련 코드(`ELEVATION_REQUIRED`, 문자열 루트 감싸기, `taskkill /T`, 남은 입력 `미실행` 기록, TOML 오류 종료 코드 2, `recheck_input`, idempotency 원자 쓰기)가 유지되고 회귀 테스트가 통과합니다. |

### 분석상의 한계

- 한글 COM의 실제 반환값·프로세스 생성 방식(예: 사용자가 연 문서가 자동화 인스턴스에 붙는지)은 버전마다 다를 수 있어 ISSUE-001의 발생 빈도를 정량화하지 못했습니다.
- 테마·토스트·위젯 등 표시 전용 UI 코드는 기능 영향이 작아 깊게 보지 않았습니다.

---

## 4. High-Risk Issues

### [ISSUE-001] 소유 한글 프로세스를 "초기화 전후 PID 차집합"으로만 판정해 사용자 한글을 앱 소유로 오인할 수 있음

- **위치:** `hwpmate/services/hwp_converter/converter.py` / `HWPConverter.initialize` (`:108` `before_pids`, `:159-162` `owned_pids = after_pids - before_pids`)
- **우선순위:** Medium
- **신뢰도:** Likely (코드 근거 명확, 실제 경쟁 재현은 하지 않음)
- **문제:** `Dispatch` 직전 스냅샷과 `RegisterModule`·`SetMessageBoxMode`·`sleep(0.2)` 이후 스냅샷의 차이에 있는 **모든** `Hwp.exe`/`HwpCtrl.exe`를 앱 소유로 봅니다. 부모 프로세스, 명령줄(COM 서버 실행 인자), 생성 시각 같은 추가 확인이 없습니다.
- **발생 조건:** 변환 시작 직후(한글 콜드 스타트 시 보통 수 초) 또는 200건마다 있는 재순환 시점에, 사용자가 탐색기에서 `.hwp`를 더블클릭하거나 한글을 직접 실행하는 경우.
- **영향:**
  - GUI는 변환 내내 `_poll_hwp_foreground`(`conversion/controller.py:144-169`, 기본 800ms 주기)로 소유 PID의 메인 창을 `hide_hwp_main_windows`로 숨깁니다. **사용자가 연 문서 창이 계속 사라집니다.**
  - 「강제 종료」(`kill_owned_processes`, 확인 창 있음)가 사용자 한글까지 종료해 **저장하지 않은 편집 내용이 사라질 수 있습니다.**
  - CLI가 이 PID를 `--owned-pids-file`에 기록하고(`app.py:416`, `:446`), MCP job이 timeout 나면 `cleanup_owned_hwp_processes`(`mcp/runner.py:247`, `jobs.py:420`)가 **확인 없이** 종료합니다.
  - 「호환 문서」 자동 응답기와 보안 창 자동 클릭도 해당 프로세스 창에 동작합니다(창 제목이 정확히 일치할 때만이라 영향은 작음).
- **근거:** `initialize`의 차집합 계산 외에 소유 여부를 다시 확인하는 코드가 없습니다. `_terminate_hwp_pids`는 이미지 이름만 다시 확인하고(PID 재사용 방지용), 같은 이름의 사용자 프로세스는 걸러내지 못합니다.
- **반증 확인:**
  - 사전 점검 창에 "변환 전 다른 한글 창을 닫는 것을 권장" 문구가 있지만, 변환 **도중** 새로 여는 경우는 막지 못합니다.
  - 이미 실행 중인 한글은 `before_pids`에 있어 제외됩니다. 문제는 그 시간 창 안에서 **새로** 시작된 프로세스뿐입니다.
  - GUI 강제 종료는 확인 창을 거치지만, 사용자는 "앱이 띄운 한글만"이라는 안내를 믿고 승인하게 됩니다.
- **호출/영향 범위:** `initialize` ← `ConversionWorker.run`, `recycle_converter`, `_run_cli_conversion`. `owned_pids` 사용처: `_suppress_hwp_ui_flash`, `_compat_dialog_responder`, `kill_owned_processes`, `engine_status_payload` → `MainWindowState.security_session` → `_poll_hwp_foreground`, CLI `_write_owned_pids` → MCP `cleanup_owned_hwp_processes`.
- **권장 수정 방향:** 차집합 후보마다 추가 확인을 하세요. (1) 부모 프로세스가 DCOM 런처(`svchost.exe -k DcomLaunch`)인지 확인합니다(`PROCESSENTRY32W.th32ParentProcessID`는 이미 구조체에 있음). (2) 명령줄에 COM 서버 실행 인자(`-Embedding`/`/Automation` 등. 한글 버전별 실측 필요)가 있는지 확인합니다. (3) 후보가 2개 이상이면 소유 집합을 비우고 추적 제한 경고로 내려가는 보수적 정책을 둡니다.
- **필요한 회귀 테스트:** `_snapshot_hwp_pids`를 `[{100}, {100, 200, 300}]`으로 monkeypatch하고 부모 PID 조회를 200→DcomLaunch, 300→explorer로 fake하면 `owned_pids == {200}`이어야 합니다. 후보 2개가 모두 DCOM이면 정책에 따라 `owned_pids == set()`이고 `process_tracking_warning`이 설정되어야 합니다.

### [ISSUE-002] 배치 도중 한글 프로세스가 죽으면 다시 연결하지 않아 다음 재순환까지 남은 파일이 모두 실패함

- **위치:** `hwpmate/workers/conversion_worker/task_runner.py` / `execute_task`, `hwpmate/workers/conversion_worker/worker.py` / `ConversionWorker.run`, `hwpmate/app.py` / `_run_cli_conversion`
- **우선순위:** Medium
- **신뢰도:** Likely (코드 흐름으로 확인, 실제 크래시는 재현하지 않음)
- **문제:** `convert_file`은 모든 COM 예외를 잡아 `(False, 메시지)`로 돌려줍니다(`converter.py:612-621`). 호출자는 실패 원인을 구분하지 않고 같은 (죽은) COM 객체로 재시도한 뒤 다음 파일로 넘어갑니다. 저장소 전체에 RPC 연결 끊김(`RPC_E_DISCONNECTED`, `RPC_S_SERVER_UNAVAILABLE`)을 감지하는 코드가 없습니다(grep으로 확인). 다시 연결되는 시점은 200건 단위 `recycle_converter`뿐입니다.
- **발생 조건:** 손상된 문서 때문에 한글이 크래시하거나, 사용자가 작업 관리자로 한글을 종료하거나, 한글이 스스로 종료되는 경우(취소 + 강제 종료는 `cancel_requested`로 루프가 멈추므로 해당 없음).
- **영향:** 크래시 시점부터 다음 200건 경계까지 최대 199개 파일이 각각 `retry_count + 1`번 실패하고 `RETRY_DELAY_SECONDS`(1초)씩 기다립니다. 결과 리포트에는 원인과 무관한 "RPC 서버를 사용할 수 없습니다" 실패가 대량으로 남습니다. 파일 손상은 없고 「실패만 재변환」으로 복구할 수 있어 Medium으로 분류했습니다. MCP는 입력마다 CLI를 새로 띄우므로 영향이 해당 입력 1건으로 제한됩니다.
- **근거:** `execute_task`의 재시도 루프(`task_runner.py:136-153`)는 `converter.convert_file`만 다시 부르고, `recycle_converter` 호출은 `worker.py`의 `(idx + 1) % CONVERTER_RECYCLE_BATCH_COUNT == 0` 조건 안에만 있습니다.
- **반증 확인:** 200건 재순환이 피해 범위를 한정해 주긴 하지만, 작은 배치(예: 150건)에서는 끝까지 복구되지 않습니다. `verify_claimed_artifacts`는 성공 판정만 다시 확인하므로 이 경우와 무관합니다.
- **호출/영향 범위:** `execute_task` ← `ConversionWorker.run`, `_run_cli_conversion`. `recycle_converter`는 이미 cleanup → 재초기화(3회 재시도)를 구현하고 있어 재사용할 수 있습니다.
- **권장 수정 방향:** 변환 실패 후 (a) `com_error.hresult`가 연결 끊김 계열이거나 (b) 소유 PID가 하나도 살아 있지 않으면, `recycle_converter`로 즉시 다시 연결하고 현재 파일을 한 번 더 시도하세요. 다시 연결하는 데 실패하면 기존처럼 `fail_pending_tasks`로 중단합니다. 반복 크래시를 일으키는 문서가 무한 재연결을 유발하지 않도록 같은 파일에서는 한 번만 다시 연결합니다.
- **필요한 회귀 테스트:** fake 엔진이 3번째 파일부터 `com_error(-2147023174)`를 던지고 재초기화 후 정상으로 돌아오게 합니다. 5건 배치에서 기대값은 성공 4~5건, 재연결 1회, 실패 메시지에 원래 오류 보존입니다. 재초기화도 실패하는 경우에는 나머지가 "재순환 실패로 중단" 메시지와 함께 실패 처리되어야 합니다.

### [ISSUE-003] 이미지/HTML 보조 산출물 판정이 형제 폴더·파일까지 포함해 출력 이름이 불필요하게 바뀜

- **위치:** `hwpmate/services/artifact_policy.py` / `is_artifact_candidate`, `matches_artifact_stem`, `iter_candidate_artifact_paths`
- **우선순위:** Low
- **신뢰도:** Confirmed (임시 폴더 재현)
- **문제:**
  1. 이미지 형식에서도 **폴더**는 `matches_artifact_stem`(구분자 `_`, `-`, `.`)으로 판정합니다. `report.hwp` 옆에 `report_2024/` 폴더가 있으면 충돌로 보고 `report (1).png`로 이름을 바꿉니다.
  2. HTML은 **파일**도 같은 규칙이라 `doc_final.html`이 있으면 `doc.html`이 `doc (1).html`이 됩니다.
  3. 일치하는 폴더 안의 파일은 확장자와 무관하게(원본 `.hwp` 포함) 최대 500개까지 스냅샷 후보에 들어갑니다.
- **발생 조건:** 덮어쓰기 꺼짐(기본값) + PNG/JPG/BMP/GIF/HTML 형식 + 같은 폴더에 `{stem}_*`, `{stem}-*`, `{stem}.*` 형태의 폴더(이미지) 또는 파일(HTML)이 있는 경우. 계획 시점과 런타임 재할당 모두 해당합니다.
- **영향:** 산출물이 `(1)` 이름으로 저장되어 사용자가 결과를 찾기 어렵습니다. 3번 때문에 변환 중 그 폴더의 파일이 다른 이유로 바뀌면 한글이 아무것도 만들지 않았어도 "변경된 산출물"로 보고 성공 처리할 수 있습니다(동시 수정이 필요하므로 가능성은 낮음). 원본 삭제는 일어나지 않습니다(§3 반증 참고).
- **근거:** 재현 스크립트 출력: `PNG conflicts: ['report_2024']`, `PNG snapshot candidates: ['notes.hwp', 'other.txt', 'report.png']`, `PNG allocated: report (1).png`, `HTML allocated: doc (1).html`, 대조군 `PDF allocated: doc.pdf`.
- **반증 확인:** 이미지 **파일**은 `_IMAGE_PAGE_SUFFIX`로 정확히 제한되어 있어 이 문제가 없고, `backup` 폴더도 제외됩니다. 문제는 폴더 판정과 HTML 파일 판정만 넓다는 점입니다. 한글 2022 실측에서 이미지 저장은 폴더를 만들지 않고(`{stem}001.png` 형식, 코드 주석), HTML은 `PIC*.png`를 같은 폴더에 만듭니다. 따라서 넓은 규칙이 필요한 근거가 코드 안에서는 보이지 않습니다.
- **호출/영향 범위:** `existing_artifact_conflicts` ← `TaskPlanner._has_existing_output_conflict` ← `allocate_output_path` ← `resolve_output_conflicts`(GUI·CLI·MCP 계획), `execute_task`(런타임). `iter_candidate_artifact_paths` ← `_snapshot_artifacts` ← `convert_file` 성공·실패 판정.
- **권장 수정 방향:** 한글 버전별 실측 산출물 규칙(이미지: `{stem}NNN.ext` 파일만, HTML: `{stem}.html` + `PIC*` + 한글이 실제로 만드는 보조 폴더 이름)을 정확히 맞추는 화이트리스트로 좁히세요. 폴더 안 중첩 후보에서는 `.hwp/.hwpx`를 제외합니다. 충돌 판정(`for_conflict=True`)에서는 폴더를 제외하는 것만으로도 1번이 해결됩니다.
- **필요한 회귀 테스트:** `report.hwp` + `report_2024/` 폴더 + PNG + overwrite=False → `report.png` 유지. `doc.hwp` + `doc_final.html` + HTML → `doc.html` 유지. `doc.hwp` + `doc001.png` + PNG → 기존처럼 `doc (1).png`(실제 페이지 충돌은 계속 감지). 스냅샷 후보에 `.hwp`가 없어야 함.

### [ISSUE-004] CLI "건너뜀만 있는" 경로에서 리포트 저장 실패가 traceback으로 끝남

- **위치:** `hwpmate/app.py` / `_run_cli_conversion` (`:390-399`)
- **우선순위:** Low
- **신뢰도:** Confirmed (재현)
- **문제:** 변환 대상이 없을 때는 `_write_cli_report`를 `try` 없이 호출합니다. 변환이 있었던 경로(`:485-490`)는 같은 실패를 `[경고] 결과 저장 실패`로 처리하고 원래 종료 코드를 유지합니다.
- **발생 조건:** 입력이 모두 목표 형식이라 건너뛰고(예: `--format HWPX`에 `.hwpx`만 있는 폴더), `--report` 경로의 부모를 만들 수 없거나 쓸 수 없는 경우.
- **영향:** 사용자·스크립트에 Python traceback이 출력되고 종료 코드 1이 됩니다. 성공(0)이어야 할 실행이 실패로 보입니다. MCP는 같은 형식 입력을 계획 단계에서 거르므로 실제 영향은 직접 CLI를 쓰는 경우에 한정됩니다.
- **근거:** 재현 결과 `FileExistsError: [WinError 183]` traceback, `exit=1`.
- **반증 확인:** `sys.excepthook = handle_exception`은 로그만 남기고(QApplication이 없어 메시지 상자도 없음) 종료 코드를 바꾸지 않습니다.
- **호출/영향 범위:** `main` → `_run_cli_conversion` → `_write_cli_report` → `atomic_io.write_results_json/csv`.
- **권장 수정 방향:** 두 경로의 리포트 저장을 같은 헬퍼로 묶어 `try/except`로 경고만 출력하세요. 종료 코드 정책(리포트 실패를 실패로 볼지)은 README에 명시합니다.
- **필요한 회귀 테스트:** `tmp_path`에 `a.hwpx`만 두고 `--format HWPX --report <파일로 막힌 경로>/r.json`으로 `main()`을 호출하면, 반환값 0(또는 정한 정책값)이고 stderr에 `결과 저장 실패`가 있으며 traceback은 없어야 합니다.

---

## 5. Potential Functional Gaps

| 항목 | 분류 | 내용 |
|------|------|------|
| 사전 점검 차단이 배치 전체를 막음 | Confirmed Gap | `PreflightDialog._blocking_errors`는 입력 1건이라도 없거나 경로가 260자 이상이면 「변환 시작」 버튼을 비활성화합니다. 1,000건 중 1건 때문에 전체를 시작할 수 없고, 해당 항목만 제외하고 진행하는 선택지가 없습니다. CLI는 사전 점검이 없어 같은 입력을 시도하므로 GUI와 CLI 동작이 다릅니다. |
| 「실패만 재변환」이 현재 출력 설정을 반영하지 않음 | Confirmed Gap (UX) | `retry_failed_tasks`는 이전 작업의 `output_file`을 그대로 쓰면서, 계획의 `same_location`/`output_path`는 현재 UI 값으로 채웁니다. 그 사이 출력 폴더를 바꾸면 사전 점검의 "저장 위치 정책"과 실제 저장 위치(상세 목록)가 서로 다르게 표시됩니다. 원래 위치에 다시 저장하는 동작 자체는 합리적입니다. |
| 업데이트 적용 시 닫기가 무시되어도 `app.quit()` 실행 | 추정 | `_apply_and_restart`는 헬퍼를 띄운 뒤 `window.close()`의 결과와 무관하게 `app.quit()`을 부릅니다. `_busy_reason`은 변환·계획만 확인하고 폴더 미리보기 스캔은 확인하지 않습니다. 스캔이 5초 안에 취소되지 않으면 `close_event`가 무시되고 설정 저장 없이 종료되며, 실행 중인 `QThread`가 파괴되면서 비정상 종료될 수 있습니다. 헬퍼는 부모 종료를 기다리므로 업데이트 자체는 진행됩니다. |
| 재순환 시 `Quit()`이 멈춘 이전 한글 | 추정 | `cleanup()`은 `Quit()` 후 바로 `owned_pids`를 비웁니다. 이전 프로세스가 끝나지 않으면 추적 대상에서 빠져 고아로 남습니다. 재순환 직후 이전 소유 PID가 살아 있는지 확인하고 필요하면 종료하는 단계가 있으면 좋습니다. |
| PrintToPDFEx 비동기 기록 | 추정 | `try_export_pdf_via_print_to_pdf_ex`는 `Execute` 직후 한 번만 산출물을 확인합니다. 프린터 드라이버가 비동기로 기록하는 환경이면 실패로 판정한 뒤 늦게 파일이 생길 수 있습니다. 실측 근거(`docs/quality/com_evidence`)는 한글 2022 환경뿐입니다. |

---

## 6. Documentation Mismatches

| 문서 | 내용 | 실제 구현 | 신뢰도 |
|------|------|----------|--------|
| `README.md` 배지 | `Tests-326 Passed` | 현재 359 passed, 1 skipped | Confirmed |
| `README.md` 개발자 가이드 | "310개 이상의 단위/통합 테스트" | 맞는 말이지만 수치가 오래되었습니다(359). | Confirmed (경미) |
| `README.md` FAQ Q7·사전 점검 문구 | "변환 전 다른 한글 창을 닫으라" | 변환 **도중** 새로 연 한글도 앱 소유로 오인될 수 있다는 점은 안내하지 않습니다(ISSUE-001). | Likely |
| `README.md` CLI 표 `--report` | "확장자로 형식 결정" | `.json`이 아니면 확장자를 `.csv`로 **바꿔서** 저장합니다(`report.txt` → `report.csv`). 문서에 명시하면 좋습니다. | Confirmed (경미) |

그 밖의 README 설명(200건 재순환, 백업 1~100개, 재시도 0~3, 단일 인스턴스 시 CLI 종료 코드 1, 원본 문서 보호, 업데이트 보류·롤백, MAX_PATH 사전 점검, Python 3.11+ MCP 요구)은 현재 구현과 일치합니다.

---

## 7. Recommended Fix Plan

### Phase 1 — Immediate

1. **ISSUE-001:** 소유 PID 판정에 부모 프로세스/명령줄 확인을 추가합니다. 확인이 애매하면 소유 집합을 비우는 보수적 정책으로 바꿉니다. MCP의 자동 종료 경로가 가장 위험하므로 `cleanup_owned_hwp_processes`에도 같은 확인을 적용합니다.
2. **ISSUE-002:** 실패 후 COM 연결 끊김을 감지하면 즉시 `recycle_converter`로 다시 연결하고 현재 파일을 한 번 더 시도합니다.

### Phase 2 — Stability

3. **ISSUE-003:** 보조 산출물 판정을 실측 규칙 화이트리스트로 좁히고, 충돌 판정에서 폴더를 빼고, 스냅샷에서 `.hwp/.hwpx`를 뺍니다.
4. **ISSUE-004:** CLI 리포트 저장 오류 처리를 하나로 묶고 종료 코드 정책을 문서화합니다.
5. 업데이트 적용 전 `_busy_reason`에 스캔 실행 여부를 넣고, `close()`가 무시되면 `quit()`을 부르지 않습니다.
6. 재순환 후 이전 소유 PID 생존 여부를 확인합니다.

### Phase 3 — Structural

7. 사전 점검 차단 항목을 "제외하고 계속"할 수 있게 하고, GUI와 CLI의 차단 정책을 맞춥니다.
8. 산출물 규칙(형식별 파일·폴더 패턴)을 `artifact_policy`의 데이터 테이블로 분리하고 한글 버전별 실측 fixture로 고정합니다.
9. `ConverterEngine` Protocol에 "연결 상태 확인"(`is_alive()` 등)을 추가해 실행기가 엔진 상태로 복구를 판단할 수 있게 합니다.

---

## 8. Test Recommendations

### Unit

- **소유 PID 판정 (ISSUE-001):** 스냅샷 `{100}` → `{100, 200, 300}`, 부모 PID fake(200: DcomLaunch, 300: explorer) → `owned_pids == {200}`. 둘 다 DcomLaunch → 정책값(빈 집합 + 경고).
- **연결 끊김 감지 (ISSUE-002):** `is_disconnect_error(com_error(hresult=-2147023174))` → True, `-2147417848`(RPC_E_DISCONNECTED) → True, 일반 SaveAs 실패 → False.
- **산출물 판정 (ISSUE-003):** §4의 4가지 입력 조합과 기대 이름. `iter_candidate_artifact_paths` 결과에 `.hwp`가 없어야 함.
- **설정 로드:** `retry_count: "2"`(문자열) → 2로 복구, `backup_max_files_per_stem: 0` → 기본값 20, JSON 손상 → `.bak`으로 옮기고 기본값 반환(기존 테스트 유지 확인).

### Integration

- **재연결 (ISSUE-002):** fake 엔진 5건, 3번째에서 끊김 → 재초기화 1회 → 최종 성공 5(또는 정책상 4)건. 재초기화 실패 → 나머지 `실패` + 중단 경고.
- **CLI 리포트 (ISSUE-004):** 건너뜀만 있는 경로 + 막힌 리포트 경로 → traceback 없음, 정한 종료 코드.
- **실패만 재변환:** 첫 실행 출력 `out1/`, 출력 폴더를 `out2/`로 바꾼 뒤 재변환 → 사전 점검의 정책 라벨과 상세 경로가 일치(또는 의도한 위치로 다시 계산).

### End-to-End (관리자 권한·한글 설치 환경, `com_real` 마커)

- 50건 배치 시작 직후 3초 안에 탐색기에서 다른 `.hwp`를 엽니다 → 사용자 창이 숨겨지지 않고, 강제 종료 대상 PID에 포함되지 않아야 합니다.
- 변환 도중 작업 관리자로 앱 소유 `Hwp.exe`를 종료합니다 → 다음 파일부터 자동으로 다시 연결해 성공해야 하며, 리포트에 재연결 경고가 1회 기록되어야 합니다.
- PNG 변환: `report.hwp` + `report_2024/` → 결과가 `report001.png…`여야 합니다.

### Concurrency

- MCP timeout 정리: 소유 PID 파일에 사용자 프로세스 PID(부모가 explorer인 fake)를 넣고 `cleanup_owned_hwp_processes` 호출 → 종료하지 않아야 합니다.
- 자동 응답 스레드 실행 중 `kill_owned_processes` 호출을 100회 반복 → 예외가 밖으로 나오지 않고 스레드가 2초 안에 멈춰야 합니다.

### Regression

- `tests/test_forbidden_patterns.py`에 "소유 PID는 차집합만으로 확정하지 않는다"를 검사하는 정적 규칙을 추가하는 것을 검토합니다(예: `owned_pids = after_pids - before_pids` 단독 대입 금지).
- 기존 P0 테스트(`test_p0_backup_isolation`, `test_p0_runtime_output_guard`, `test_p0_crosspath_parity`)를 산출물 규칙 변경 후에도 그대로 유지합니다.

### Platform-specific (Windows)

- 경로 255~259자 입력: 사전 점검 경고만 표시하고 변환은 확장 경로(`\\?\`)로 성공해야 합니다. 260자 이상은 차단(또는 Phase 3 이후 "제외하고 계속").
- UNC 경로(`\\server\share\a.hwp`): `com_path_candidates`가 `\\?\UNC\…` 후보를 만들고 백업 폴더 생성이 가능한지 확인합니다.
- cp949 콘솔에서 CLI 실행: 이모지 출력이 `?`로 바뀌고 종료 코드가 정상이어야 합니다(기존 수정의 회귀 확인).

---

## 9. Final Assessment

| 항목 | 평가 | 근거 |
|------|------|------|
| Functional Correctness | **Good** | 계획 → 런타임 재할당 → 산출물 검증 → PDF 매직 검사가 GUI·CLI 공통으로 일관됩니다. 남은 오류는 보조 산출물 판정의 오탐(Low)입니다. |
| Runtime Stability | **Acceptable** | 취소·강제 종료·종료 흐름이 재진입까지 고려되어 있습니다. 다만 한글 크래시 후 다시 연결하지 않습니다(ISSUE-002). |
| Data Integrity | **Good** | 원본 덮어쓰기 금지, `Clear(1)`, "이전에 있던 파일은 지우지 않음", 원자적 쓰기가 지켜집니다. 사용자 한글 프로세스 오인(ISSUE-001)만 예외입니다. |
| Error Resilience | **Acceptable** | 대부분의 외부 실패(설정 손상, DLL 잠금, 업데이트 교체·스모크 실패, TOML 오류)를 복구합니다. COM 연결 끊김과 CLI 일부 경로가 빠져 있습니다. |
| Cross-platform Robustness | **Acceptable** | Windows 전용 제품으로서 긴 경로·UNC·cp949·권한 상승을 다룹니다. Windows 외 OS는 설계상 대상이 아니며, MCP도 job을 실행할 때 `PLATFORM_UNSUPPORTED`로 거부합니다(`mcp/jobs.py:52`, `:382`). |
| Test Confidence | **Good** | 359개 테스트, pyright 0, P0 회귀 고정, 금지 패턴 정적 검사가 있습니다. 실제 COM·크래시·동시 실행 시나리오는 opt-in(`com_real`)으로만 존재합니다. |

### 실제로 먼저 수정할 문제 3개

1. **[ISSUE-001]** 소유 한글 PID 판정 강화. 사용자 문서 창 숨김과 확인 없는 MCP 자동 종료를 막습니다.
2. **[ISSUE-002]** COM 연결 끊김 감지 후 즉시 다시 연결. 한글 크래시 한 번이 최대 199건 실패로 번지는 것을 막습니다.
3. **[ISSUE-003]** 이미지/HTML 보조 산출물 판정 축소. 불필요한 `(1)` 이름 변경과 스냅샷 오염을 없앱니다.

---

## 10. 수정 결과 (2026-11-07)

§4~§6의 모든 항목을 수정했습니다. 실제 한글 2022 환경에서 검증하는 과정에서 감사 때 놓친 버그 1건(아래 NEW-001)을 추가로 찾아 함께 고쳤습니다.

### 이슈별 조치

| # | 조치 | 주요 위치 | 검증 |
|---|------|----------|------|
| ISSUE-001 | 차집합 후보마다 부모 프로세스를 확인합니다. COM LocalServer는 DcomLaunch 서비스(`svchost.exe`)가 띄우므로, 부모가 svchost이거나 이미 채택된 한글인 PID만 소유로 확정합니다. 부모를 확인할 수 없는 후보는 후보가 하나뿐일 때만 채택합니다(기존 동작 유지). 시작 실패 시 고아 정리와 MCP timeout 정리(`_live_hwp_pids`)는 부모 불명도 제외해 더 보수적으로 동작합니다. | `process_snapshot.select_com_launched_pids`, `converter.initialize`, `mcp/runner._live_hwp_pids` | 단위 테스트 8개, 정적 규칙(`owned_pids = a - b` 금지). 실측: 이 PC의 LocalServer32는 `hwp.exe -Automation`이고, Dispatch로 뜬 `hwp.exe`의 부모는 `svchost.exe`로 채택됨 |
| ISSUE-002 | `convert_file` 예외에서 연결 끊김 HRESULT(`RPC_S_SERVER_UNAVAILABLE`, `RPC_E_DISCONNECTED` 등 5종)를 감지하고, `is_alive()`(끊김 플래그 + 소유 PID 생존)를 추가했습니다. `execute_task`는 실패 후 엔진이 죽었으면 작업당 한 번 즉시 재연결해 같은 파일을 다시 시도합니다(재시도 횟수는 쓰지 않음). 재연결에 실패하면 `EngineUnavailableError`로 남은 작업을 실패 처리하고 중단합니다. GUI는 새 소유 PID를 다시 알리고, CLI는 `--owned-pids-file`을 다시 씁니다. | `converter.is_com_disconnect_error`/`is_alive`, `task_runner.execute_task`, `worker.run`, `app._run_cli_conversion` | 단위·통합 테스트 9개. **실측:** 4건 배치 중 소유 `Hwp.exe`를 `taskkill`로 종료 → `-2147023174` 감지 → 재연결 → 4건 모두 성공, 끝난 뒤 남은 `Hwp.exe` 없음 |
| ISSUE-003 | 이미지 형식은 폴더를 산출물로 보지 않습니다. HTML은 같은 stem의 파일(`doc.htm` 등)과 `{stem}_files`·`{stem}.files` 폴더만 봅니다. 폴더 안 스냅샷에서는 `.hwp/.hwpx`를 뺍니다. | `artifact_policy.is_artifact_candidate`, `iter_candidate_artifact_paths` | 감사 재현 2건을 회귀 테스트로 고정. 실측: 텍스트 문서에서 HTML은 `doc.html`만, PNG는 `doc001.png`·`doc002.png`만 생성되고 폴더는 생기지 않음 |
| ISSUE-004 | 두 경로의 리포트 저장을 `_save_cli_report` 하나로 묶었습니다. 실패하면 `[경고] 결과 저장 실패`만 출력하고 종료 코드는 변환 결과로 정합니다. README에 이 정책을 적었습니다. | `app._save_cli_report` | 감사 재현 조건(파일로 막힌 경로)으로 반환값 0, traceback 없음 |
| **NEW-001** | **CLI의 200건 재순환과 재연결이 항상 실패하던 버그.** CLI는 `initialize(manage_com_apartment=True)`로 컨버터가 COM apartment를 소유하는데, `recycle_converter`가 `cleanup()`(CoUninitialize) 뒤 `initialize(manage_com_apartment=False)`를 불러 "CoInitialize가 호출되지 않았습니다"로 실패했습니다. 이제 원래 소유 여부를 유지해 다시 엽니다. GUI는 워커 스레드가 apartment를 소유하므로 영향이 없었습니다. | `task_runner.recycle_converter` | 회귀 테스트 1개. 실측: 수정 전 재연결 3회 모두 실패 → 수정 후 재연결·주기적 재순환 모두 성공 |

### 기능 공백(§5) 조치

| 항목 | 조치 |
|------|------|
| 사전 점검 차단이 배치 전체를 막음 | 차단 사유를 작업별로 기록하고, 일부만 차단되면 **「문제 N개 제외하고 시작」** 버튼을 보여 줍니다. 제외한 항목은 결과에 「건너뜀」과 사유로 남습니다. 출력 폴더 쓰기 불가는 그 폴더의 모든 작업에 적용합니다. |
| 「실패만 재변환」 라벨 불일치 | `PlannedConversion.output_policy_note`를 추가했습니다. 재변환 계획은 "이전 변환과 같은 출력 경로 (현재 출력 폴더 설정은 적용하지 않음)"으로 표시합니다. |
| 업데이트 적용 시 닫기가 무시되어도 `app.quit()` 실행 | `_busy_reason`이 폴더 스캔 실행도 확인합니다. `window.close()`가 `False`를 돌려주면 `quit()`하지 않고 다시 시도하라고 안내합니다. |
| 재순환 시 `Quit()`이 멈춘 이전 한글 | `cleanup()`이 Quit 뒤 이전 소유 PID가 최대 3초 안에 끝나는지 확인하고, 남아 있으면 PID 지정으로 종료합니다. |
| PrintToPDFEx 비동기 기록 | `Execute`/`RunToPDF`가 성공을 반환했는데 파일이 아직 없으면 최대 3초 기다리고, 파일 크기가 안정됐는지 확인한 뒤 판정합니다. 실패를 반환하면 기다리지 않습니다. |

### 문서(§6) 조치

README의 테스트 배지(326 → 390)와 개발자 가이드 수치, `--report` 확장자 규칙(`.json` 외에는 `.csv`로 바꿈), 리포트 저장 실패 시 종료 코드 정책, 변환 중 사용자가 연 한글은 추적 대상이 아니라는 안내(Q7), 즉시 재연결, 사전 점검 「제외하고 시작」, 「실패만 재변환」 저장 위치를 반영했습니다.

### 검증

- `pytest`: **390 passed, 1 skipped**(기본 설정. `com_real` 등 opt-in 제외). 수정 전 359개 + 신규 31개.
- `pyright`: 0 errors.
- 실제 한글 2022 COM(관리자 권한 아님, 스크래치 폴더의 합성 문서): PID 부모 판정, HTML/PNG 산출물 규칙, 크래시 후 재연결, CLI 경로 재순환을 확인했습니다.
- **확인하지 못한 것:** GUI 화면에서 「제외하고 시작」 버튼 조작(오프스크린 Qt 테스트로만 확인), 실제 업데이트 교체, 한글 2018/2020 등 다른 버전의 부모 프로세스·산출물 규칙, 그림이 들어간 문서의 HTML `PIC*` 산출물.
