# HwpMate MCP 확장 설계 및 구현 지시서

> **대상:** `twbeatles/HwpMate` (**HwpMaster가 아님**)  
> **상태:** 제안 설계 / 구현 전 / 기존 저장소 변경 없음  
> **기준일:** 2026-10-08 (KST)  
> **분석 기준 커밋:** [`542bfa42838697d9c0a666ce1f4b8a4f1fb7833f`](https://github.com/twbeatles/HwpMate/tree/542bfa42838697d9c0a666ce1f4b8a4f1fb7833f)  
> **확인 버전:** README 및 `hwpmate/constants.py` 표기 **v9.1.2**  
> **문서 목적:** 이미 작동하는 HWP/HWPX 일괄 변환 GUI·CLI를 보존하면서, Codex·Claude Code 등 에이전트가 명시적 권한 하에 변환 작업을 계획·요청·추적·결과 검증할 수 있는 로컬 MCP 서버를 추가한다.

## 0. 에이전트 최상위 구현 지침

1. `HwpMate`의 **기존 CLI·변환 엔진·TaskPlanner를 재작성하지 않는다.** MCP는 orchestration/adapter 계층이다.
2. 초기 MCP는 **Windows + 한컴오피스 한글 + pywin32** 환경의 기능을 노출한다. macOS/Linux 변환 지원을 약속하지 않는다.
3. **문서 변환은 읽기 전용 동작이 아니다.** 출력 파일, 원본 백업, report, 임시 파일이 생성되므로 변환 도구에는 쓰기권한·확인 절차가 필요하다.
4. 기존 `SingleInstanceLock`이 GUI/CLI의 COM 동시 실행을 배제한다. MCP는 자체 Job Queue로 중복 호출을 막고, 기존 잠금을 **우회하거나 제거하지 않는다**.
5. 변환은 기본적으로 **별도 HwpMate CLI subprocess**에서 실행한다. MCP 서버 프로세스가 COM 객체를 직접 보유하지 않는다. 향후 안정화 후에만 직접 서비스 호출 방식 검토.
6. 초기에는 **MCP `stdio`만 제공**하고 원격 HTTP 호출, 무인 자동 승인, 임의 `.exe` 실행, 임의 명령 문자열 입력을 금지한다.
7. 기존 자동 백업·원본 HWP/HWPX 보호·출력 충돌 회피·재시도·산출물 검증·프로세스 재순환 정책을 유지한다.
8. 기본 변환 요청은 `preview → submit → status/result` 흐름으로 설계한다. `--overwrite`, `--no-backup`, 임의 경로는 MCP에 처음부터 노출하지 않는다.
9. 모든 예시 코드·명령은 **구현 설계용**이다. 기존 CLI와 신규 MCP 인터페이스를 명확히 구분한다.

---

## 1. 코드베이스 실사

### 1.1 현재 배포·사용 모델

- HwpMate는 PyQt6 GUI와 헤드리스 CLI를 가진 Windows 한글 문서 일괄 변환기다.
- `.hwp`와 `.hwpx`를 입력받고 `PDF`, `DOCX`, `HWPX`, `HWP`, `ODT`, `HTML`, `RTF`, `TXT`, `PNG`, `JPG`, `BMP`, `GIF` 등 `FORMAT_TYPES`에 열거된 형식으로 변환한다.
- 기존 `hwpmate/app.py`는 `argparse` CLI를 이미 제공하며, `--input`, `--format`, `--output`, `--recursive`, `--overwrite`, `--no-backup`, `--retry`, `--pdf-export-mode`, `--report`, `--no-auto-continue`를 인식한다.
- `_run_cli_conversion()`은 `TaskPlanner.build_tasks()`, `resolve_output_conflicts()`, `HWPConverter.initialize(manage_com_apartment=True)`, `execute_task()`, `recycle_converter()`를 사용한다.
- `ConversionSummary`와 `ConversionTask`가 결과 모델을 정의하고 `write_results_json()` / `write_results_csv()`로 결과 리포트를 생성한다.
- 변환에 앞서 기존 `SingleInstanceLock`을 획득하며, 잠금 실패 시 CLI 종료 코드 1을 반환한다.
- README는 200건 단위 COM 인스턴스 재순환, PDF 듀얼 엔진, 원본 자동 백업 및 `.hwp/.hwpx` 원본 덮어쓰기 방지를 설명한다.

### 1.2 실제 재사용 모듈 및 주의점

| 소스 | 확인된 역할 | MCP 적용 |
|---|---|---|
| `hwpmate/app.py` | argparse, `_run_cli_conversion`, `--report`, CLI 엔트리 | **P0/P1 subprocess가 재사용할 기존 실행 경로** |
| `hwpmate/services/task_planner.py` | `TaskPlanner.build_tasks`, `resolve_output_conflicts`, `allocate_output_path` | read-only preflight 계획 및 충돌 미리보기 |
| `hwpmate/models.py` | `ConversionTask`, `PlannedConversion`, `ConversionSummary.to_json_dict()` | 결과 DTO와 report parsing |
| `hwpmate/services/artifact_policy.py` | 원본 확장자 보호, 멀티파일 보조 산출물 검출 | 전체 산출물 정책 적용 |
| `hwpmate/services/hwp_converter/` | HWP COM 변환 엔진 | MCP 내부에서 직접 중복 구현하지 않음 |
| `hwpmate/workers/conversion_worker/task_runner.py` | `TaskRunOptions`, `execute_task`, 재순환 | 기존 CLI 경로에서 그대로 사용 |
| `hwpmate/ui/dialogs/atomic_io.py` | 결과 CSV/JSON 원자적 저장 | `--report <job>.json` 결과 신뢰 기반 |
| `hwpmate/app_instance.py` | `QLockFile`, `%LOCALAPPDATA%/HwpMate/HwpMate.lock` | GUI/CLI 동시 실행 제외 |
| `hwpmate/services/hwp_print_settings/` | PDF 저장 모드 정규화 | `saveas_first`·`print_to_pdf_ex_first` |
| `tests/test_cli_conversion.py`, `test_task_planner.py`, `test_app_instance.py` | 이미 존재하는 CLI/계획/잠금 테스트 | MCP 회귀 테스트 기준선 |

#### 특히 중요한 기술적 사실

1. `hwpmate/app.py`가 파일 상단에서 `PyQt6.QtWidgets`와 GUI `MainWindow`를 import한다. **MCP 서버 진입점에서 `hwpmate.app`을 import하지 않는 구조**가 처음에는 더 안정적이다. 기존 CLI를 별도 프로세스로 실행한다.
2. `--smoke`는 import 및 업데이트 모듈 검사와 `PYWIN32_AVAILABLE` 등의 정보를 제공하지만, **한글 COM 실제 변환 성공을 보증하지 않는다**. `hwp_capabilities` 및 선택적 COM 현장 검증을 별도로 정의한다.
3. CLI는 화면용 진행 메시지를 stdout에 출력한다. **stdout을 MCP JSON-RPC 채널로 사용하는 같은 프로세스 안에서 CLI를 호출하면 프로토콜이 깨진다.** subprocess stdout/stderr를 파일/파이프로 수집하고 MCP stdout으로 전달하지 않는다.
4. `_run_cli_conversion()`이 `--report`로 만든 JSON 파일은 결과 파싱용으로 활용할 수 있다. **CLI 종료 코드만으로 개별 파일 성공을 판정하지 않는다.**
5. CLI는 `--output` 디렉터리가 없으면 생성할 수 있고, `TaskPlanner.build_tasks()`는 시점에 따라 미존재 디렉터리를 거부할 수 있다. 미리보기 시 출력 폴더 준비/생성의 책임을 명확히 구분한다.
6. `TaskPlanner`는 출력 충돌 시 새로운 파일명을 할당한다. 계획과 실행 사이 파일이 생길 수 있는 **TOCTOU 경합**이 있으므로 preview의 최종 출력 파일명은 절대 보장되지 않는다.
7. DOCX·RTF 저장의 '계속' 확인 창 자동 처리, 보안 팝업 승인 등은 앱이 시작한 한글 프로세스에서만 동작한다는 정책을 유지한다. **MCP가 임의의 외부 한글 창을 조작해서는 안 된다.**

### 1.3 기존 CLI 실제 명령 (현행)

```powershell
# 현행 예시: 실제 CLI 인터페이스
HwpMate-v9.1.2.exe --input "C:\Docs\report.hwp" --format PDF
HwpMate-v9.1.2.exe --input "C:\Docs" --format DOCX --recursive --output "C:\Exports"
HwpMate-v9.1.2.exe --input "C:\Docs" --format PDF --report "C:\Reports\result.json"
```

HwpMate EXE가 없는 소스 개발 환경에서는 프로젝트의 공식 런처를 확인하고 Python 모듈 기반 호출 경로를 사용한다. 아래 MCP 설계에서는 실제 **절대 경로로 지정한 승인된 실행 바이너리**만 호출한다.

---

## 2. MCP 제품 정의

### 2.1 목표

**HwpMate MCP = AI 에이전트가 HWP 파일 변환을 계획하고 안전하게 배치 실행하며, 완료 상태와 산출물 검증 결과를 받는 도구**

대표 요청:
- "이 디렉터리의 HWP 파일 중 PDF로 변환 가능한 파일을 목록으로 보여줘."
- "이 승인된 문서 3개를 PDF로 변환해줘. 원본은 보존해줘."
- "변환에 실패한 문서만 이유와 함께 정리해줘."
- "변환이 끝난 보고서의 결과 파일 경로, 생성 파일 수, 경고를 알려줘."

### 2.2 비목표 및 경계

- HwpMate MCP는 **문서 편집·자연어 내용 작성·표 변형·개인정보 자동 제거 도구가 아니다.** 그것은 HwpMaster/HwpOps 계열의 영역이다.
- `rhwp` 도입, 크로스플랫폼 문서 엔진 교체, Windows COM 대체는 이 설계서의 범위가 아니다.
- 로컬 관리자 권한 획득을 자동화하지 않는다. 기존 GUI의 관리자 권한 요구와 CLI 동작을 혼동하지 말고 실제 COM 필요 권한을 별도로 검증한다.
- 소스 파일 삭제, 원본 덮어쓰기, 임의 스크립트/쉘 실행, HWP 편집은 제공하지 않는다.
- 파일 본문을 MCP로 읽어 LLM에 전송하는 기능은 P0에서 제공하지 않는다.

### 2.3 구현 전략 비교

| 전략 | 장점 | 리스크 | 권고 |
|---|---|---|---|
| 기존 CLI를 subprocess로 호출 | 성숙한 변환 파이프라인 재사용, COM 격리, 현재 CLI 회귀 위험 적음 | 진행률 구조화 및 취소 제어 어려움 | **P0/P1 우선** |
| 직접 `HWPConverter`를 MCP worker에서 호출 | 진척률·취소·개별 파일 제어 쉬움 | COM apartment·GUI/worker 수명·잠금 재설계 위험 | P2 이후 검토 |
| MCP가 임의 CLI 명령 실행 | 구현 단순 | 명령주입·보안정책 우회 | **금지** |
| 장기 실행 원격 HTTP MCP | 여러 클라이언트 연결 | 인증·한글 COM 세션·멀티유저 격리 복잡 | 초기 제외 |

---

## 3. 제안 아키텍처

```text
Codex / Claude Code
       |
   MCP stdio  ← JSON-RPC stdout 전용
       |
 hwpmate.mcp.server
       |
 ┌─────┴─────────────────────┐
 |                           |
Plan / Preflight        Job Service
(read-only)                 |
  TaskPlanner          allowlist / queue
                           |
                    subprocess launcher
                           |
                     HwpMate CLI
                     --input --format
                     --output --report
                           |
                  기존 SingleInstanceLock
                           |
                       한컴 COM
                           |
                output + report.json + logs
                           |
                    Job Result Parser
                           |
                   MCP structured result
```

- MCP 프로세스는 변환 중 상태·job ID를 관리하되 한글 COM 객체를 소유하지 않는다.
- **동시 변환 최대 1개**: 기존 GUI나 다른 CLI가 실행 중이면 `BUSY_EXTERNAL_INSTANCE`를 반환하거나 설정된 대기 정책에 따라 제한적 재시도. lock 우회 금지.
- 내부 큐 기본 최대 3개 대기, 대기 만료 10분 등 수치는 초기 제안값이므로 실제 부하 테스트 후 설정으로 확정.
- MCP stdio 서버는 사용자 권한으로만 실행. 승격된 GUI 프로세스와 비승격 MCP 프로세스가 같은 `SingleInstanceLock`을 정확히 공유하는지 Windows 실기에서 검증.
- 긴 작업은 MCP 호출을 열린 상태로 무한 대기하지 않고 `job_id`를 반환한다.
- Job store는 `%LOCALAPPDATA%/HwpMate/mcp/jobs/` 등 권한 제한 디렉터리를 사용한다. 폴더 생성 시 owner-only ACL을 적용하고 오래된 로그·보고서의 보존 기간을 지정한다.

### 3.1 제안 파일 배치

```text
hwpmate/mcp/
    __init__.py
    __main__.py              # python -m hwpmate.mcp
    server.py                # MCP 선언과 stdio 진입점
    schemas.py               # Pydantic input/output / job state
    config.py                # 허용 루트·경로·quota·로그 정책
    capabilities.py          # 플랫폼/실행파일/pywin32 등 preflight
    policy.py                # 경로와 포맷 허용, overwrite 금지
    planner_adapter.py       # TaskPlanner를 통한 read-only planning
    jobs.py                  # 단일 작업 큐 / idempotency / recovery
    runner.py                # 승인된 CLI subprocess 실행
    report_reader.py         # 기존 ConversionSummary JSON 매핑
    errors.py                # 일관된 에러 코드
    logging_utils.py        # stderr/file log only

tests/mcp/
    test_stdio_protocol.py
    test_capabilities.py
    test_plan.py
    test_paths_and_acl.py
    test_job_lifecycle.py
    test_cli_subprocess.py
    test_report_mapping.py
    test_lock_contention.py
    test_cli_gui_regression.py
```

### 3.2 패키징 정책

- 현행 루트에 `pyproject.toml`이 없는 배치를 확인했다. 프로젝트의 `requirements.txt`, PyInstaller spec과 배포 방식을 먼저 검토한다.
- 최소 설치 예: MCP 추가 의존성을 `requirements-mcp.txt`에 분리하거나 `[project.optional-dependencies]` 체계를 도입하되 기존 GUI 실행과 frozen 배포를 깨지 않는다.
- 소스 설치 진입: **신규 제안 명령** `python -m hwpmate.mcp`.
- Windows 배포 EXE에 MCP를 넣을지, CLI/GUI EXE와 MCP bridge EXE를 분리할지는 결정 사항으로 남긴다. **권장: MCP bridge는 별도 콘솔 바이너리로 패키징**, 변환 실행은 기존 HwpMate EXE 절대 경로를 호출.
- PyInstaller `--windowed` GUI EXE를 MCP 자체 서버로 겸용하지 않는다. stdout, 서명 검증, 관리자 권한 정책을 별도로 시험한다.

---

## 4. MCP 도구 설계

**다음 도구는 모두 신규 제안이다. 기존 HwpMate CLI에 현재 존재한다고 표기하지 않는다.** 도구 이름 충돌을 막기 위해 `hwpmate_` 접두사를 고정한다.

| 도구 | 성격 | 입력 | 주요 출력 | 단계 |
|---|---|---|---|---|
| `hwpmate_get_capabilities` | 읽기 | 없음 | OS, HwpMate 버전, CLI 경로 유효성, 포맷, 환경 상태 | P0 |
| `hwpmate_list_supported_formats` | 읽기 | 없음 | `FORMAT_TYPES` 기반 포맷 목록·주의사항 | P0 |
| `hwpmate_preview_conversion` | 읽기(파일시스템 조사) | `input_paths`, `format`, `output_dir`, `recursive=false` | 개수, 목표 경로, 충돌·건너뜀·제약, `plan_id` | P0 |
| `hwpmate_submit_conversion` | 쓰기 | `plan_id`, `confirmation`, `idempotency_key` | `job_id`, 초기 상태 | P1 |
| `hwpmate_get_job_status` | 읽기 | `job_id` | `queued/running/...`, 성공/실패 개수 | P1 |
| `hwpmate_get_job_result` | 읽기 | `job_id`, `limit`, `cursor` | 개별 산출물·오류·경고 | P1 |
| `hwpmate_cancel_job` | 변경 | `job_id`, `confirmation` | 취소 요청 수락 및 상태 | P2 (보수적) |
| `hwpmate_validate_artifacts` | 읽기/검증 | `job_id` | 파일 존재·크기·타입·검증 상태 | P2 |

**도구 수는 일부러 적게 유지한다.** 개별 변환용과 배치 변환용을 나눠 중복하지 말고 `input_paths` 배열로 공통 처리한다. 단, 첫 버전에서 기존 CLI가 `--input` 하나만 받으므로 **여러 입력은 파일별 순차 CLI 작업 또는 사전 준비된 허용 폴더 모드로 변환**해야 한다. 절대로 사용자 입력을 문자열 결합해 셸 명령으로 만들지 않는다.

### 4.1 `hwpmate_get_capabilities`

반환 내용:

```json
{
  "schema_version": "hwpmate-mcp/v1",
  "platform": "windows",
  "app_version": "9.1.2",
  "cli_path_valid": true,
  "pywin32_available": true,
  "hancom_com_probe": "not_run",
  "supported_formats": ["PDF", "DOCX", "HWPX", "HWP", "TXT"],
  "instance_lock": "unknown",
  "warnings": ["실제 COM 변환 가능 여부는 현장 테스트 필요"]
}
```

이것은 **표현 형식 예시**다. `supported_formats`는 위처럼 일부만 하드코딩하지 말고 `FORMAT_TYPES` 전부에서 생성한다. `hancom_com_probe`는 `available/unavailable/not_run` 중 하나로 표현하되 단순 `--smoke` 결과로 `available`을 선언하지 않는다.

### 4.2 `hwpmate_preview_conversion` 계약

```json
{
  "input_paths": ["C:\\Work\\Reports"],
  "format": "PDF",
  "output_dir": "C:\\Work\\Exports",
  "recursive": false
}
```

예상 응답:

```json
{
  "schema_version": "hwpmate-mcp/v1",
  "plan_id": "opaque-plan-handle",
  "expires_at": "2026-10-08T10:20:00+09:00",
  "format": "PDF",
  "requested": 5,
  "planned": 4,
  "skipped": 1,
  "conflicts_renamed": 1,
  "requires_confirmation": true,
  "warnings": ["변환 계획 시점과 실행 시점 사이 파일명이 바뀔 수 있습니다"],
  "preview": []
}
```

- `TaskPlanner.build_tasks()`와 `resolve_output_conflicts()`의 계획 정보를 활용한다. **미리보기 중 실제 문서 변환 금지**.
- `input_paths`는 파일·폴더 혼합을 허용하는 공통 인터페이스지만, 내부적으로 여러 planner 호출을 합산·중복제거한다. 파일/폴더 혼합이 불명확할 경우 P0에서는 하나의 파일 또는 하나의 폴더만 허용해도 된다.
- 파일 수·총 byte 크기·깊이 제한을 적용한다. 예: 기본 최대 50개, 500MB, 재귀 depth 5 (가이드 값, 운영 측정 후 조정).
- 숨김 폴더/`backup` 폴더/출력 폴더 재귀 재수집 방지.
- 미리보기 시 생성될 `plan_id`는 비추측형 handle이며 입력 옵션, 파일 stat/mtime/size, 출력 정책·포맷, 계획 해시를 묶는다. 만료 전에도 입력이 변경되면 `PLAN_STALE`을 반환한다.
- 실행 직전 동일 정책으로 경로 및 충돌을 **재검증**한다. 대상 경로 재할당이 생기면 별도 승인 갱신 또는 원본 미변경 원칙을 지키며 보고한다.

### 4.3 `hwpmate_submit_conversion` 계약

```json
{
  "plan_id": "opaque-plan-handle",
  "confirmation": "approve_non_destructive_conversion",
  "idempotency_key": "user-supplied-retry-key"
}
```

성공 시 `job_id`, `status=queued|running`, 예상 개수, 제출 시각만 반환한다. `confirmation`은 **서버 정책상의 명시적 승인 플래그**일 뿐, MCP 클라이언트의 사용자 확인을 대체하지 않는다. 호스트 측 도구 승인 UI 사용 여부를 추가로 검증한다.

초기 기본값 고정:

- `overwrite=false`.
- `backup_enabled=true`.
- `retry_count=1` (기존 CLI 기본값, 설정으로 상한 0~3).
- `pdf_export_mode=saveas_first`.
- `auto_continue_compat_dialog` 정책은 기존 HwpMate 동작을 유지하되 관리자가 활성/비활성 설정을 고정할 수 있게 한다. 임의 원격 에이전트가 허가 창 정책을 변경할 수 없도록 한다.
- 기존 `.hwp/.hwpx` 원본 보호 로직은 어떤 옵션에서도 우회 불가.
- 출력 폴더는 사용자 설정의 **승인된 output root 아래**에만 둔다. 입력 근처 자동 저장은 MCP에서 기본 금지.

### 4.4 작업 상태 머신

```text
planned --submit--> queued --> running --> succeeded
                         |         |--> partially_failed
                         |         |--> failed
                         |         |--> cancel_requested --> canceled / interrupted
                         |         |--> timeout / interrupted
                         |--> rejected_busy / expired
```

- `job_id`는 **서버 관리용 작업 단위**이고, HwpMate CLI에는 별도 job ID 기능이 없다.
- 같은 `idempotency_key` + 같은 계획은 **동일 job ID**를 재반환한다. 동일 key로 다른 계획 제출 시 `IDEMPOTENCY_CONFLICT`.
- `succeeded`는 CLI return code 0 **그리고** 유효한 report.json **그리고** 보고서상 실패·취소가 0건일 때에만 선언한다.
- 중간 오류·종료 코드만 0·report 누락/손상 등은 실패/검증 필요 상태로 처리한다. 보고서 검증 실패를 성공으로 보정하지 않는다.
- MCP 서버 재시작으로 작업 관리자 상태가 사라져도 저장된 job record가 있으면 `interrupted` 또는 `recovery_required`로 복원한다. **파일 성공 여부가 불명확한 작업을 자동 재실행하지 않는다.**

---

## 5. 안전한 실행 Runner 구현

### 5.1 CLI subprocess(권장)

호출 빌더는 허용된 필드에서 고정 인자 배열만 만든다.

```python
# 의사 코드: 실제 호출은 검증된 absolute executable path와 args를 사용
argv = [
    str(approved_hwp_mate_exe),
    "--input", str(approved_input),
    "--format", output_format,
    "--output", str(approved_output_dir),
    "--report", str(job_private_dir / "result.json"),
    "--retry", "1",
]
# subprocess.Popen(argv, shell=False, cwd=..., env=sanitized_env, ...)
```

- Python source 모드는 허용된 인터프리터 + 공식 진입점으로 실행할 수 있다. **사용자가 `command` 필드를 제공하는 API는 금지**.
- `shell=False`, `cwd` 고정, 환경변수 최소화, stdout/stderr 캡처 길이 제한, 프로세스·작업 제한 시간.
- 실행 파일 해시/버전·서명 검증은 배포 형태에 맞게 선택한다. 신뢰되지 않은 PATH 검색 결과로 실행파일을 교체하지 않는다.
- `--report`는 job별 **고유한 JSON 경로**를 사용한다. 변환 결과를 stdout 텍스트 정규식으로 파싱하지 않는다.
- 파일명이 한글·공백·괄호여도 인자 배열로 전달하여 셸 이스케이프 실수를 피한다.
- stdout/stderr의 변환 로그가 MCP 프레임으로 새어나가지 않아야 한다.
- 여러 파일 변환이 필요하면 입력 각 파일을 직렬 subprocess로 수행하거나 별도 내부 배치 entry를 신규 도입한다. **한글 COM 동시 worker 여러 개는 시작하지 않는다.**

### 5.2 잠금 및 병렬성

현재 `SingleInstanceLock`은 `QLockFile`을 사용하며 기본 경로는 `%LOCALAPPDATA%/HwpMate/HwpMate.lock`이다. `setStaleLockTime(30_000)`이 설정돼 있으므로 **오래 걸리는 정상 변환을 stale로 오판하는지**를 다중 프로세스 Windows 실기에서 확인해야 한다.

- MCP 프로세스 내 job queue도 1개 worker만 활성화한다.
- CLI subprocess가 실제 기존 잠금을 잡게 한다. MCP가 잠금 파일을 직접 삭제하거나 lock을 강제 해제해서는 안 된다.
- GUI 실행 중 submit 시 `BUSY_EXTERNAL_INSTANCE`를 반환하거나 명시적으로 큐 대기 처리한다.
- 취소·timeout 시 무작정 `taskkill /IM hwp.exe /F` 등 전체 한글 프로세스를 종료하지 않는다. 현재 CLI가 시작·소유한 프로세스만 종료 가능한지 명확히 증명되기 전까지 자동 hard kill을 금지한다.
- HwpMate 업데이트 진행 중 작업 제출 불가, 변환 중 업데이트 적용 보류라는 기존 정책을 유지한다.

### 5.3 취소 및 오류 복구

현재 CLI 루프는 `cancel_check=lambda: False`로 task를 실행하고, Ctrl+C는 상위 계층에서 취급한다. 따라서 MCP `cancel`을 단순 이벤트 플래그로 구현하면 **실제 취소되지 않는다**.

초기 권장안:
- P1: 실행 전 대기 중인 작업만 `cancel` 가능. 실행 중은 `CANCELLATION_UNSUPPORTED_WHILE_RUNNING`을 반환해 정직하게 안내.
- P2: `execute_task`의 cancel signal을 CLI까지 연결하고, 안전한 COM 작업 종료·정리·보고서 생성이 검증되면 실행 중 취소 도입.
- CLI가 치명적 COM 오류로 종료한 경우 job 로그·report 여부를 보존해 재처리/수동 확인 가능하도록 한다.

---

## 6. 보안 모델

### 6.1 기본 정책 예시

```toml
# %LOCALAPPDATA%/HwpMate/mcp/config.toml (신규 제안)
[mcp]
transport = "stdio"
enabled = true
max_parallel_conversions = 1
max_queued_jobs = 3
max_files_per_job = 50
max_input_total_mb = 500
max_log_kb = 256
require_preview = true
require_confirmation = true
allow_overwrite = false
allow_disable_backup = false
allow_remote_http = false
allow_unc_paths = false
allow_reparse_points = false
allow_running_cancel = false

[paths]
input_roots = ["C:/Work/HWP-Input"]
output_roots = ["C:/Work/HWP-Output"]
cli_executable = "C:/Program Files/HwpMate/HwpMate-v9.1.2.exe"
```

- 설정의 경로는 **샘플**이며 사용자가 허용 디렉터리로 직접 수정해야 한다. **루트 설정이 없으면 변환 도구는 fail closed**.
- 경로는 `Path.resolve()` 후 root containment를 검사한다. Windows drive-letter 대소문자, `..`, junction/reparse point, symlink, UNC, `\\?\` 확장 경로 등 공격 사례를 테스트한다.
- Windows reparse point는 `resolve()`만으로 완전한 보호가 보장되지 않는다. 경로 확인 후 실행 직전 재검증, 파일 핸들 기반 검사 등 OS 수준 추가 점검을 설계한다.
- 모델이 임의 경로를 제시하더라도 기존 UI의 최근 폴더·설정 전체를 그대로 공개하지 않는다.
- 하위 출력 디렉터리 생성은 명시적인 변경 작업으로 취급하고, 승인된 output root 밖으로 이동하지 않게 한다.
- XML/HWP의 내용은 모델에서 반환할 필요가 없으므로, 기본 MCP 응답에는 문서 본문을 포함하지 않는다.
- 로그·리포트에 파일명이 포함될 수 있으므로 접근 권한 제한, 보존 기간 및 민감값 마스킹을 적용한다.
- 문서 내용을 외부 클라우드 LLM이 읽게 하는 확장 도구는 별도 사용자 동의·데이터 반출 검토 후 구현한다.

### 6.2 도구 annotation 제안

| 도구 | annotation 제안 | 이유 |
|---|---|---|
| get_capabilities / list_formats | `readOnlyHint=true` | 환경·메타데이터 조회 |
| preview_conversion | `readOnlyHint=true` | 문서 변환/파일 생성 없음이 실증된 경우에 한함 |
| submit_conversion | `readOnlyHint=false`, `destructiveHint=false`, `idempotentHint=false` | 신규 파일·백업·로그 생성; 완전한 멱등은 key 강제 전 불가 |
| get_job_status / get_job_result | `readOnlyHint=true` | 저장된 상태 읽기 |
| cancel_job | `readOnlyHint=false`, `destructiveHint=true` | 작업 상태 및 일부 산출물에 영향 가능 |

annotation은 **보안 통제가 아니라 모델/호스트에 제공하는 힌트**이다. 구체 권한·확인·경로 제한을 서버가 강제해야 한다.

### 6.3 오류 코드 및 대응

| 코드 | 의미 | 처리 |
|---|---|---|
| `PLATFORM_UNSUPPORTED` | Windows 아님 | 변환 비활성 |
| `HANCOM_UNAVAILABLE` | 한글 COM 초기화 불가 | 사용자에게 설치/구동 확인 요청 |
| `CLI_NOT_TRUSTED` | 실행파일 경로·검증 실패 | 실행 차단 |
| `INPUT_DENIED` / `OUTPUT_DENIED` | root 바깥 경로 | 거부 |
| `BAD_FORMAT` | `FORMAT_TYPES` 미지원 | 허용 포맷 안내 |
| `PLAN_STALE` / `PLAN_EXPIRED` | 계획 변경/만료 | 다시 preview 요청 |
| `CONFIRMATION_REQUIRED` | 명시적 승인 없음 | 실행 금지 |
| `BUSY_EXTERNAL_INSTANCE` | GUI/다른 CLI 실행 중 | 대기/나중에 재요청 |
| `JOB_QUEUE_FULL` | 대기 한도 초과 | 새 job 거부 |
| `REPORT_MISSING` | CLI report 미생성 | 성공 주장 금지 |
| `OUTPUT_VALIDATION_FAILED` | 산출물 검증 불통과 | 실패/검토 필요 |
| `CANCELLATION_UNSUPPORTED_WHILE_RUNNING` | 안전 취소 미구현 | 실행 중 강제 종료 금지 |
| `TIMEOUT` / `INTERNAL` | 제한시간/예외 | 상태 보존 후 원인 조사 |

---

## 7. MCP Resources / Prompts

### Resources

- `hwpmate://capabilities` : 현재 서버의 변환 가능 환경(한컴 COM 미실행 probe를 구분)
- `hwpmate://formats` : 출력 포맷·특징
- `hwpmate://jobs/{job_id}` : **현재 사용자가 제출한 job** 요약; 예측 불가능한 ID와 소유자 검증 필요

**금지:** `file://` 경로를 그대로 MCP resource로 공개해 임의 로컬 파일 다운로드를 허용하지 않는다.

### Prompts

- `convert_hwp_to_pdf_safely`: 입력 폴더 확인 → preflight → 사용자 승인 → job submit → 검증.
- `prepare_documents_for_office`: 포맷 요구만 분석하고 HwpMate가 지원하는 범위의 변환 계획을 제시.
- `review_conversion_failures`: job 결과·실패 사유를 요약하고 자동 무한 재시도 없이 재검토.

---

## 8. 단계별 구현 로드맵

### Phase 0 — 기준선 및 현장 환경 확인

- [ ] 분석 기준 커밋과 현재 HEAD 차이를 확인하고, GUI/CLI를 사용해 기존 변환 테스트 기록.
- [ ] `python -m pytest -q`, Pyright·기존 빌드 절차 등 프로젝트 기준선 실행.
- [ ] 실제 Windows 한컴 버전, 관리자 권한 없이 CLI 사용 가능 여부, 멀티 사용자 잠금 정책 조사.
- [ ] CLI `--report` JSON schema, stdout/stderr 인코딩, 오류 종료 코드, 기존 테스트에 추가 fixture 생성.
- [ ] 공식 Python MCP SDK v2와 Codex/Claude Code 지원 조합 검증하고 의존성 버전 고정.
- **완료:** 기존 GUI/CLI 변환·원본 보호·기본 배치 테스트 결과 확보.

### Phase 1 — read-only 계획 API

- [ ] `hwpmate/mcp/config.py`로 허용 input/output roots와 quota 구현.
- [ ] `hwpmate/mcp/planner_adapter.py`에 `TaskPlanner` 재사용. `preview` 호출에서 COM 미시작·파일 생성 없음 보장.
- [ ] `hwpmate_get_capabilities`, `hwpmate_list_supported_formats`, `hwpmate_preview_conversion` 구현.
- [ ] multi-input/recursive 제한을 명확히 정하고 입력 중복 제거·path containment·reparse point 방어.
- [ ] `stdout` 오염 없음/서버 정상 initialize/tools/list 테스트.
- **완료:** HWP 파일을 변환하지 않고 에이전트가 변환 가능성·위험·출력 예정 경로를 조회.

### Phase 2 — 기존 CLI subprocess 기반 P1 Job Runner

- [ ] 실행파일 경로 명시·검증·고정, 명령 인자 배열 생성, `shell=False`.
- [ ] `JobStore` 구현: plan hash, owner/session, idempotency, timestamps, state, report path, log path.
- [ ] 동시 작업 1개·대기 한도·상태 조회·재시작 recovery·TTL cleanup.
- [ ] `hwpmate_submit_conversion`, `hwpmate_get_job_status`, `hwpmate_get_job_result` 구현.
- [ ] CLI `--report <job>.json` 활용, 결과 스키마 검증 및 실패 시 보수적 상태 전환.
- [ ] lock 경쟁, 동시 GUI 실행, 출력 파일 충돌, PDF 2가지 export 모드 검증.
- **완료:** 사람이 승인한 HWP→PDF·HWP→DOCX 작업을 안전하게 단일 worker에서 수행.

### Phase 3 — 실제 산출물 검증 및 취소

- [ ] `hwpmate_validate_artifacts`: 경로·크기·확장자·서명(PDF `%PDF-`, DOCX ZIP 등)·복수 이미지 파일 목록 확인.
- [ ] PDF 열기/페이지 수 검증은 선택형; 원본과의 시각 동일성은 **보증하지 않는다**.
- [ ] 실행 중 취소는 COM task runner의 협조적 cancel·프로세스 소유권 식별이 테스트되기 전까지 비활성.
- [ ] timeout 후 남는 owned process의 정리 정책을 Windows 실제 환경에서 검증.
- **완료:** 무인 CLI 실패에 대해 원인·부분 성공·실패 파일을 재현할 수 있음.

### Phase 4 — 배포 및 클라이언트 호환성

- [ ] 신규 stdio MCP 런처를 Windows용 콘솔 실행파일로 패키징하거나 source install 지원.
- [ ] 기존 `HwpMate-v9.1.2.exe`·GUI install의 업데이트 서명/롤백 정책 유지.
- [ ] `docs/MCP.md`, `docs/MCP_SECURITY.md`, `docs/MCP_TOOLS.md`, `docs/MCP_CLIENT_SETUP.md` 작성.
- [ ] Windows 10/11 + 한글 버전별 smoke 및 Codex/Claude Code 실제 연결 테스트.
- [ ] 공용 오픈소스로 배포한다면 내부자료·경로 로그·실험 문서를 제거하고 최소 권한 기본값 확인.

---

## 9. 필수 테스트 매트릭스

| 시나리오 | 조건 | 기대 결과 |
|---|---|---|
| stdio 독립성 | MCP 서버 켠 상태 | stdout에 JSON-RPC 외 GUI/CLI 문구 없음 |
| 포맷 일치 | 모든 `FORMAT_TYPES` | MCP 포맷 목록과 기존 앱 일치 |
| 한컴 미설치 | pywin32/COM 불가능 | preview 가능 여부 구분, submit 거부, 무한 대기 없음 |
| GUI 실행 중 | HwpMate GUI가 lock 보유 | BUSY, 동일 COM 작업 중복 없음 |
| 2개 submit 동시 | 계획 2건 | 활성 worker 1개, 큐 정책 준수 |
| 파일 접근 제한 | `C:\Windows`, `..`, symlink/junction, UNC | 차단 또는 명시적 허용만 |
| 입력 변경 | preview 후 입력 교체/삭제 | `PLAN_STALE`, 변환 금지 |
| 원본 보존 | HWP/HWPX 이름 충돌 | 원본 hash 불변, 출력 이름 변경 |
| 출력 보호 | 기존 PDF와 충돌 | 기본 덮어쓰기 없이 새 이름 사용 |
| backup | 정상 변환 | 원본 백업 정책 유지 |
| HTML·이미지 | 보조 자산 생성 | 결과에 모든 생성물 포함, 누락 식별 |
| PDF export | 2종 모드 | 기존 폴백·유효성 검사 유지 |
| COM hang | 대기/timeout | `TIMEOUT`, 임의 `hwp.exe` 전체 종료하지 않음 |
| 중간 실패 | N개 중 일부 실패 | `partially_failed` + 파일별 결과 |
| report 누락 | CLI 비정상 조기 종료 | `REPORT_MISSING`, 성공으로 표시 안 함 |
| JSON report | 한글·특수문자 파일 | 한글 비손상·스키마 검증 통과 |
| 재호출 | 동일 idempotency key | 중복 변환 0회, 동일 job 반환 |
| 재시작 | running 중 MCP 종료 | 재시작 후 `interrupted`/복구 경고 |
| 실제 Windows 설치 | Win 10/11, 한컴 설치 | GUI/CLI/MCP 회귀 테스트 수행 |
| 보안 로그 | 민감 파일명·경로 | 접근 통제/마스킹/보존기간 준수 |

### 테스트 명령

```powershell
# 기존 테스트 (현행 구조에서 동작 여부 먼저 확인)
python -m pytest -q
python -m pytest tests/test_cli_conversion.py tests/test_task_planner.py tests/test_app_instance.py -q

# 신규 MCP 구현 후 기대 테스트
python -m pytest tests/mcp -q
python -m hwpmate.mcp --help
```

새 테스트에서는 실제 COM 통합 테스트를 별도 mark로 두고, CI에서 pywin32/한컴 없는 환경에서도 mock 기반 MCP contract 테스트가 통과하도록 한다. **실제 COM 검증 없이 모든 기능이 작동한다고 주장하지 않는다.**

---

## 10. MCP 구성 예시 (구현 이후 유효)

```json
{
  "mcpServers": {
    "hwpmate": {
      "command": "C:\\Python\\python.exe",
      "args": ["-m", "hwpmate.mcp"]
    }
  }
}
```

- 위는 `mcpServers` 형식을 쓰는 클라이언트용 **예시 설정**이다. 경로는 설치 환경에 맞게 바꿔야 한다.
- Codex는 별도 TOML 설정 구조를 사용하므로 버전별 공식 설정에 맞게 `docs/MCP_CLIENT_SETUP.md`에 제공한다.
- `python -m hwpmate.mcp`는 **신규 제안 진입점**이며 현재는 존재한다고 가정하지 않는다.
- MCP가 파일 변환 후 PDF 내용을 직접 모델에 반환하지 않는다. 반환하는 것은 검증 결과·파일 경로·건수·오류와 경고다.

---

## 11. 에이전트에 붙여넣는 구현 프롬프트

> `twbeatles/HwpMate`를 확인하고 이 문서의 기준 커밋 이후 HEAD 변경을 우선 조사하라. 기존 CLI `_run_cli_conversion`, `TaskPlanner`, `ConversionSummary`, `SingleInstanceLock`, `artifact_policy`, 변환 worker와 테스트를 재사용하여 HwpMate에 로컬 stdio MCP를 추가한다. HwpMaster와 혼동하지 않는다. P0는 capabilities/formats/preview 세 read-only 도구만 구현하고, 그다음 P1에서 submit/status/result를 추가한다. MCP는 한컴 COM 객체를 직접 생성하지 않고 **검증된 기존 HwpMate CLI subprocess를 `shell=False`로 호출**한다. 단일 활성 worker, 명시적 input/output root allowlist, preview·approval·idempotency, 기존 `.hwp/.hwpx` 보호, 백업 on, overwrite off를 구현한다. GUI/CLI의 기존 잠금을 우회하지 않는다. stdin/stdout은 MCP 프로토콜 전용이며 변환 로그는 별도 파일·stderr로 분리한다. `--report` JSON을 검증해 job 성공을 판정하고, timeout·lock 경쟁·재실행·충돌·부분 실패를 mock/Windows COM 실기 테스트로 확인하라. GUI/CLI 회귀가 없음을 입증하고, 실제 구현 범위·미구현 위험·시험 결과를 문서에 기록하라.

---

## 12. 출시 전 완료 조건 (Definition of Done)

- [ ] `HwpMate` 현행 CLI/GUI 및 자동 업데이트에 회귀가 없다.
- [ ] MCP `stdio` 도구 발견과 구조화 반환이 실제 클라이언트에서 작동한다.
- [ ] `preview`는 실제 변환과 산출물 파일 생성 없이 수행된다.
- [ ] default policy에서 임의 폴더 접근, 원본 덮어쓰기, 백업 해제가 불가능하다.
- [ ] COM 동시 실행은 최대 1개로 제한되며 GUI/CLI lock을 우회하지 않는다.
- [ ] 작업 ID와 상태 조회로 장시간 변환과 재시작을 다룰 수 있다.
- [ ] CLI 결과 JSON의 성공/실패/건너뜀/취소와 MCP 결과가 일치한다.
- [ ] 출력 보조 파일을 포함한 산출물 검증, 부분실패, report 누락 처리가 검증된다.
- [ ] Windows 실기/한컴 설치 환경에서 최소 HWP→PDF 및 HWP→DOCX 통합 테스트 통과.
- [ ] 사용법·권한·보안·장애복구·클라이언트 설정 문서 작성.

## 13. 참고 자료

- [HwpMate README](https://github.com/twbeatles/HwpMate/blob/542bfa42838697d9c0a666ce1f4b8a4f1fb7833f/README.md)
- [현재 CLI 구현](https://github.com/twbeatles/HwpMate/blob/542bfa42838697d9c0a666ce1f4b8a4f1fb7833f/hwpmate/app.py)
- [TaskPlanner](https://github.com/twbeatles/HwpMate/blob/542bfa42838697d9c0a666ce1f4b8a4f1fb7833f/hwpmate/services/task_planner.py)
- [ConversionSummary 모델](https://github.com/twbeatles/HwpMate/blob/542bfa42838697d9c0a666ce1f4b8a4f1fb7833f/hwpmate/models.py)
- [SingleInstanceLock](https://github.com/twbeatles/HwpMate/blob/542bfa42838697d9c0a666ce1f4b8a4f1fb7833f/hwpmate/app_instance.py)
- [산출물 보호 정책](https://github.com/twbeatles/HwpMate/blob/542bfa42838697d9c0a666ce1f4b8a4f1fb7833f/hwpmate/services/artifact_policy.py)
- [결과 원자적 기록](https://github.com/twbeatles/HwpMate/blob/542bfa42838697d9c0a666ce1f4b8a4f1fb7833f/hwpmate/ui/dialogs/atomic_io.py)
- [MCP Python SDK v2](https://py.sdk.modelcontextprotocol.io/)
- [MCP tools 명세(2026-07-28)](https://github.com/modelcontextprotocol/modelcontextprotocol/blob/main/docs/specification/2026-07-28/server/tools.mdx)
- [MCP annotation 참고](https://blog.modelcontextprotocol.io/posts/2026-03-16-tool-annotations/)

**적용 전 필수:** HEAD·지원 SDK 버전·Windows 한컴 실행 환경을 재검증할 것. 이 문서는 실제 MCP 구현 결과가 아니라 구현 지시서다.
