# HwpMate MCP 도구 레퍼런스

스키마 버전: `hwpmate-mcp/v1`. 실패 시 공통 envelope은
`{schema_version, ok: false, error: {code, message}}`다.

## hwpmate_get_capabilities (읽기)

입력 없음. OS·앱 버전·CLI 경로 유효성·포맷 목록·환경 상태를 반환한다.
`hancom_com_probe`는 항상 `not_run`이며, `--smoke`만으로 COM 가능을 선언하지 않는다.

## hwpmate_list_supported_formats (읽기)

`FORMAT_TYPES` 12종 전체를 `{name, ext, desc}`로 반환한다.

## hwpmate_preview_conversion (읽기)

입력: `{input_paths: string[], format = "PDF", output_dir: string, recursive = false}`.
실제 변환·파일 생성 없이 계획과 `plan_id`(10분 TTL)를 반환한다.

출력: `{plan_id, expires_at, format, requested, planned, skipped,
conflicts_renamed, requires_confirmation: true, warnings, preview: [{input, output}]}`.
`preview`는 최대 100건이다. 출력 폴더가 없어도 미리보기는 폴더를 만들지 않으며
제출 시 생성됨을 경고한다.

## hwpmate_submit_conversion (쓰기)

입력: `{plan_id, confirmation: "approve_non_destructive_conversion", idempotency_key}`.
출력: `{job_id, state}`. 같은 key+계획은 동일 `job_id`를 재반환하고,
같은 key+다른 계획은 `IDEMPOTENCY_CONFLICT`다.

## hwpmate_get_job_status (읽기)

`{job_id}` → `{state, requested, success/failed/skipped/canceled_count, ...}`.
상태: `queued → running → succeeded | partially_failed | failed | timeout |
canceled | interrupted | rejected_busy | expired`.
대기 큐에서 `queue_wait_timeout_seconds`를 넘기면 실행 없이 `expired`가 된다.

## hwpmate_get_job_result (읽기)

`{job_id, limit = 50 (최대 200), cursor = 0}` → `{total, items, next_cursor, warnings, error}`.
`items[] = {input, outputs, status, detail, retry_count}`.

## hwpmate_cancel_job (변경)

`{job_id, confirmation}` → `{job_id, state}`. 대기 중만 취소되며,
실행 중은 `CANCELLATION_UNSUPPORTED_WHILE_RUNNING`이다.

## hwpmate_validate_artifacts (읽기)

`{job_id}` → `{checked, verified, invalid: [{path, reasons}], all_verified}`.
성공 항목의 존재·크기·확장자·매직 서명(PDF `%PDF-`, DOCX/ZIP `PK..`, HWP OLE 등)을
확인한다. 시각 동일성은 보증하지 않는다.

## Resources / Prompts

- `hwpmate://capabilities`, `hwpmate://formats`, `hwpmate://jobs/{job_id}`
- `convert_hwp_to_pdf_safely`, `prepare_documents_for_office`, `review_conversion_failures`
