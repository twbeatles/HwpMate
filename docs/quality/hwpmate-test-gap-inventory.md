# HwpMate 테스트 갭 인벤토리

- 기준 커밋: `833f653d94c7d77a0ad15d2795e51490dea128db` (종료 시점 동일, 작업물 uncommitted)
- 측정: `pytest --cov=hwpmate --cov-branch` 전체 68% (line+branch)
- 범례: `보강됨` = 이번 검증에서 신규 테스트·수정으로 커버, `잔여` = 남은 위험·후속 과제

## P0 — 원본·백업·산출물·교차 경로

| 영역 | 상태 | 근거 |
|---|---|---|
| 동일 stem `.hwp`/`.hwpx` 원본 보호 (계획+실행 시점) | 보강됨 | 기존 planner 테스트 + 신규 `test_p0_runtime_output_guard.py` 6건, mutation M1/M3로 검출 입증 |
| 배치 중복 출력 분리·내용 비섞임 | 보강됨 | 신규 배치 테스트 + 실 COM에서 `sample_ko (1).pdf/.docx` 자동 개명 실측 |
| 계획 후 출력 선점(TOCTOU) 무음 교체 금지 | 보강됨 | 신규 TOCTOU 테스트 + `execute_task` 변환 직전 재할당 (기존 동작 회귀 고정) |
| 엔진 성공 보고 후 산출물 실재 검증 | 보강됨 | `verify_claimed_artifacts` 신규 + 테스트 2건 + mutation M4. 단, 산출물을 보고하지 않는 엔진은 기존 신뢰 유지(의도적) |
| summary 집계 vs 파일시스템 실재 대조 | 보강됨 | `test_summary_counts_match_files_on_disk` |
| 백업 stem 격리·타임스탬프 정렬·실패 기록 | 보강됨 | `test_p0_backup_isolation.py` 4건 + mutation M2. 백업 실패는 변환 계속 + task·경고에 명시(현행 정책 유지) |
| GUI/CLI/MCP 계획 parity | 보강됨 | `test_p0_crosspath_parity.py` 3건 (planner vs preview 분류·출력명 일치) |
| 산출물 내용 검증(PDF 페이지·DOCX ZIP·이미지 디코드) | 보강됨 | 실 COM 5종 + pypdf/PIL/zip 독립 오라클 대조 (증거 `docs/quality/com_evidence/`) |
| cp949 콘솔 CLI 충돌 | 보강됨(수정) | 실측 `UnicodeEncodeError` → `_ensure_cli_printable_output` + 회귀 테스트 |

## P1 — 배치·MCP·업데이트

| 영역 | 상태 | 근거 |
|---|---|---|
| COM 재순환·취소·worker 수명 | 기존 커버 유지 | `test_converter_recycle.py`, worker 취소·재시도 테스트 통과 확인 |
| MCP 승인·경로·멱등·큐·PLAN_STALE | 기존 커버 + 검출 입증 | 기존 60+ 테스트 + mutation M5/M6/M8 |
| MCP validator 내용 불신 | 보강됨 | `test_p1_mcp_validator.py` 11건(52%→82%) + live `test_validate_distrust.py` 2건 |
| MCP 설정 로딩 | 보강됨 | `test_p1_mcp_config.py` 7건(49%→91%) |
| 서명·업데이트·롤백 | 기존 커버 유지 | manifest/installer/apply-script 테스트 통과 확인 (위조·만료·다운그레이드·롤백) |
| 단일 인스턴스 락·설정 복구 | 기존 커버 유지 | `test_app_instance.py`, config 테스트 통과 확인 |

## 잔여 갭 (후속 승인 필요)

1. `hwpmate/services/hwp_converter/converter.py` 62% — COM 분기(PrintToPDFEx 폴백·보안모듈 실패·재시도)는 실 COM에서만 실행됨. hermetic 추가는 fake 비대화 위험으로 보류.
2. `hwpmate/services/hwp_security_module.py` 40%, `windows_integration/*` 31~86% — 관리자 권한·실 레지스트리 영역. `release_vm` 마커로 분리 필요.
3. `hwpmate/workers/file_scan_worker.py` 33%, GUI 컨트롤러 41~68% — 2초+ 스캔 차단 등 과거 감사 항목은 코드상 캐시(`folder_file_paths`)로 완화되어 있으나 장시간 스캔 E2E는 미실행.
4. 500건+ 실제 COM 스트레스는 미실행 (시간·리소스 승인 필요). fake 200-­배치 재순환은 기존 테스트로 커버.
5. `pytest.ini`에 `com_real`/`gui_real`/`release_vm` 마커 등록 + 기본 제외. 실 COM E2E는 QA 스크립트 증거로 대체하고 커밋 테스트에는 넣지 않음.
