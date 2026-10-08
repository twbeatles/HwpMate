# HwpMate MCP 보안 모델

## 원칙

- 문서 변환은 읽기 전용이 아니다(출력·백업·report·로그 생성). 변환 제출에는
  명시적 승인 토큰(`approve_non_destructive_conversion`)과 `idempotency_key`가 필요하다.
- 도구 annotation(`readOnlyHint` 등)은 모델·호스트용 힌트이며 보안 통제가 아니다.
  경로 제한·승인·충돌 회피는 서버가 강제한다.
- 기본값: `overwrite=false`, `backup=true`, `retry=1`, `pdf_export_mode=saveas_first`.
  원본 `.hwp/.hwpx` 보호는 어떤 옵션에서도 우회할 수 없다.
- 초기에는 MCP `stdio`만 제공하며, 원격 HTTP 호출·무인 자동 승인·임의 `.exe` 실행·
  임의 명령 문자열 입력을 금지한다.

## 경로 정책

- 모든 경로는 `resolve()` 후 허용 루트 containment를 검사한다.
  루트 미설정 시 변환 도구는 fail closed.
- `..` 탈출·드라이브 밖 경로·UNC(기본 거부)·`\\?\` 확장 경로·symlink/junction
  (기본 거부)은 차단하고, 실행 직전 계획을 재검증한다(`PLAN_STALE`).
- 출력은 승인된 output root 아래에만 둔다. 입력 근처 자동 저장은 기본 금지이므로
  미리보기·제출에는 출력 폴더 지정이 필수다.

## 동시성·잠금

- MCP 내부 job 큐는 단일 worker만 활성화한다.
- 기존 `SingleInstanceLock`을 우회·삭제하지 않는다. GUI/다른 CLI 실행 중 제출은
  `BUSY_EXTERNAL_INSTANCE`로 보고하고, 소유하지 않은 한글 프로세스를 종료하지 않는다.
- 실행 중 취소는 지원하지 않는다(`CANCELLATION_UNSUPPORTED_WHILE_RUNNING`).
  대기 중 작업만 취소할 수 있다.

## 데이터 취급

- 기본 MCP 응답에는 문서 본문을 포함하지 않는다(경로·건수·오류·경고만 반환).
- job 기록·로그·보고서는 `%LOCALAPPDATA%\HwpMate\mcp\jobs\`에 보관되며,
  파일명에 민감 정보가 포함될 수 있어 접근 권한을 제한한다.
  종료된 기록은 `job_retention_days`(기본 7일) 후 정리된다.

## 오류 코드

| 코드 | 의미 |
|---|---|
| `PLATFORM_UNSUPPORTED` | Windows가 아님 (예약) |
| `HANCOM_UNAVAILABLE` | 한글 COM 초기화 불가 (현장 확인 필요) |
| `CLI_NOT_TRUSTED` | 실행 파일 경로·검증 실패 |
| `INPUT_DENIED` / `OUTPUT_DENIED` | 허용 범위 밖 경로·UNC·reparse |
| `BAD_INPUT` / `BAD_FORMAT` | 한글 문서 아님·미지원 형식 |
| `INPUT_TOO_LARGE` | 파일 수·총 크기 quota 초과 |
| `PLAN_STALE` / `PLAN_EXPIRED` | 입력 변경·계획 만료 |
| `CONFIRMATION_REQUIRED` | 명시적 승인 없음 |
| `BUSY_EXTERNAL_INSTANCE` | GUI/다른 CLI 실행 중 |
| `JOB_QUEUE_FULL` / `JOB_NOT_FOUND` / `JOB_NOT_DONE` | 큐·조회 상태 |
| `IDEMPOTENCY_CONFLICT` | 같은 key로 다른 계획 제출 |
| `REPORT_MISSING` | CLI 보고서 미생성·손상 (성공으로 보정 금지) |
| `OUTPUT_VALIDATION_FAILED` | 산출물 검증 불통과 |
| `CANCELLATION_UNSUPPORTED_WHILE_RUNNING` | 실행 중 취소 불가 |
| `TIMEOUT` / `INTERRUPTED` / `INTERNAL` | 제한 시간·큐 만료·재시작 중단·내부 오류 |
| `HANCOM_UNAVAILABLE` / `PLATFORM_UNSUPPORTED` | 한글 COM 실패 매핑·비Windows 실행 차단 |
