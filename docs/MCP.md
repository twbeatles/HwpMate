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

설정 파일: `%LOCALAPPDATA%\HwpMate\HwpMate\mcp\config.toml` — 예시:

```toml
[mcp]
transport = "stdio"
enabled = true
max_queued_jobs = 3
max_files_per_job = 50

[paths]
input_roots = ["C:/Work/HWP-Input"]
output_roots = ["C:/Work/HWP-Output"]
cli_executable = "C:/Program Files/HwpMate/HwpMate-v9.2.0.exe"
```

- 허용 루트가 없으면 변환 도구는 fail closed(거부)한다.
- 환경 변수 `HWPMATE_MCP_INPUT_ROOTS`·`HWPMATE_MCP_OUTPUT_ROOTS`·`HWPMATE_MCP_CLI`
  로도 지정할 수 있다.

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
