# HwpMate MCP 서버

AI 에이전트(Codex·Claude Code 등)가 명시적 승인 하에 HWP/HWPX 일괄 변환을
계획·요청·추적·검증할 수 있는 로컬 MCP 서버(`stdio` 전용)다.
대상 저장소는 `twbeatles/HwpMate`이며, 기존 GUI·CLI·변환 엔진은 그대로 재사용한다.

## 설치

```powershell
pip install -r requirements.txt
pip install -r requirements-mcp.txt
```

## 실행

```powershell
# 소스 설치 실행 (신규 진입점)
python -m hwpmate.mcp
```

- `transport`는 `stdio`만 지원한다. 원격 HTTP·무인 자동 승인은 제공하지 않는다.
- MCP 프로세스는 한컴 COM 객체를 직접 보유하지 않고, 검증된 HwpMate CLI를
  `shell=False` subprocess로 호출한다.
- stdin/stdout은 MCP JSON-RPC 전용이다. 변환 로그는 job별 로그 파일로만 기록된다.

## 설정

설정 파일: `%LOCALAPPDATA%\HwpMate\mcp\config.toml` (`--config <경로>`로 변경 가능) — 예시:

```toml
[mcp]
transport = "stdio"
enabled = true
max_queued_jobs = 3
max_files_per_job = 50
job_timeout_seconds = 3600   # 입력 파일 1개(CLI 1회)당 제한 시간

[paths]
# 반드시 목록([...])으로 적는다. 드라이브 전체("C:/")는 허용 루트로 쓸 수 없다.
input_roots = ["C:/Work/HWP-Input"]
output_roots = ["C:/Work/HWP-Output"]
# 소스 진입점 권장 (아래 '실행 권한' 참고)
cli_executable = "C:/Tools/HwpMate/hwptopdf-hwpx_v4.py"
```

- 허용 루트가 없으면 변환 도구는 fail closed(거부)한다.
- 환경 변수 `HWPMATE_MCP_INPUT_ROOTS`·`HWPMATE_MCP_OUTPUT_ROOTS`(`;` 구분)·`HWPMATE_MCP_CLI`
  로도 지정할 수 있다.
- 설정 파일 문법·형식이 잘못되면 서버는 시작하지 않고 stderr에 원인을 한 줄로 남긴다
  (종료 코드 2). Windows 경로는 `"C:/Work"`처럼 `/`를 쓰거나 `'C:\Work'`처럼 작은따옴표로 감싼다.
- 설정 파일을 쓰려면 Python 3.11 이상이 필요하다(`tomllib`). 3.10에서는 설정 파일이 있으면
  시작을 거부한다.
- `allow_overwrite`·`allow_disable_backup`·`allow_running_cancel`·`require_preview`·
  `require_confirmation`·`allow_remote_http`·`max_parallel_conversions`는 안전상 고정값으로만
  동작한다. 다른 값을 적으면 시작 로그와 `hwpmate_get_capabilities` 경고에 표시된다.

### 실행 권한 (중요)

배포 exe(`HwpMate-v*.exe`)는 관리자 권한 매니페스트(`requireAdministrator`)로 빌드된다.
일반 권한으로 실행된 MCP 호스트(Claude Code·Codex 등)는 이 exe를 실행할 수 없으며
(`WinError 740`), job은 `ELEVATION_REQUIRED`로 실패한다. 다음 중 하나를 사용한다.

1. `cli_executable`에 소스 진입점 `hwptopdf-hwpx_v4.py`를 지정한다
   (MCP 서버와 같은 Python으로 실행됨, 권장).
2. MCP 호스트 자체를 관리자 권한으로 실행한 뒤 exe를 지정한다.

`hwpmate_get_capabilities`의 `cli_requires_elevation`이 `true`이면 현재 권한으로 실행할 수 없는 상태다.

## 기본 흐름

```text
preview → (사용자 승인) → submit → status/result → validate
```

1. `hwpmate_preview_conversion`으로 계획·충돌·`plan_id` 확보 (읽기 전용).
2. 사용자에게 입력·출력·건수를 보여주고 승인 획득.
3. `hwpmate_submit_conversion`(`confirmation="approve_non_destructive_conversion"`,
   `idempotency_key` 필수)으로 제출.
4. `hwpmate_get_job_status`·`hwpmate_get_job_result`로 추적.
5. `hwpmate_validate_artifacts`로 산출물 검증.

상세: [MCP_TOOLS.md](MCP_TOOLS.md), [MCP_SECURITY.md](MCP_SECURITY.md),
[클라이언트 설정](MCP_CLIENT_SETUP.md).
