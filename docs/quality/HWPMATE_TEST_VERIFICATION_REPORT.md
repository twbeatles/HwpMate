# HwpMate 테스트 검증 결과

- 기준 commit: `833f653d94c7d77a0ad15d2795e51490dea128db` / 종료 commit: 동일 (작업물 uncommitted)
- OS·Python: Windows 11 (ARM64) / 프로젝트 `.venv` Python 3.11.9 (시스템 3.14에서 `cryptography` 빌드 실패 → 3.11 venv로 격리, `cryptography`는 wheel-only 설치)
- HwpMate 버전: v9.2.0 / 한글 버전: 12.0.0.4605 (ProgID `HWPFrame.HwpObject`, 비상승 `--allow-non-admin`)
- 작업 경로: 리포 내 `.qa-scratch/` (합성 전용 input/output/reports/jobs, 검증 후 삭제, 증거 JSON만 `docs/quality/com_evidence/`에 보존)
- 원본 해시 manifest: `docs/quality/com_evidence/fixture_manifest.json` (합성 90줄, 개인정보 없음)

## 기존·변경 후 전체 pytest / pyright

| 시점 | pytest | pyright | 비고 |
|---|---|---|---|
| 변경 전 baseline | 282 passed, 1 skipped (54s) | clean (venv 인터프리터 지정 시) | skip 1건 = symlink 권한 부족(기존 사유) |
| 변경 후 최종 | 316 passed, 1 skipped (64s) | 0 errors, 0 warnings | 신규 34건 포함, 동일 skip 1건 |
| 커버리지 | 전체 67% → 68% (line+branch) | — | validator 52%→82%, mcp config 49%→91%, task_runner 88%→91%, backup 88%→92% |

- 모의 테스트 범위: 전체 316건 (hermetic, `QT_QPA_PLATFORM=offscreen`, `PYTHONUTF8=1`, CI와 동일 조건)
- 실 COM/GUI/CLI/MCP E2E 범위: 아래 별도 기록. 모의 통과를 현장 성공으로 주장하지 않음.

## TEST-ID (신규)

| ID | 시나리오 / oracle | 호출 경로 | 결과 |
|---|---|---|---|
| P0-G1 | 계획 후 출력 선점 시 무음 교체 금지·원본 바이트 불변 | `execute_task` + fake 엔진 | 통과 |
| P0-G2 | `overwrite=true`でも 기존 `.hwp`/`.hwpx` 위 기록 금지·SHA-256 불변 | `execute_task` (planner 우회 조건) | 통과 |
| P0-G3 | 동일 출력 요구 2건 분리 저장·내용 비섞임 | `execute_task` 공유 키 | 통과 |
| P0-G4 | 성공 보고 + 산출물 없음 → 실패 | `verify_claimed_artifacts` (신규) | 통과 |
| P0-G5 | 성공 보고 + 0바이트 → 실패 | 동상 | 통과 |
| P0-G6 | summary 집계 == 디스크 실재 파일 | `build_summary` 대조 | 통과 |
| P0-B1 | 타 stem 백업 개수·해시 불변 | `create_backup` 정리 | 통과 |
| P0-B2 | 백업 정렬은 파일명 타임스탬프 (mtime 아님) | 동상 | 통과 |
| P0-B3/B4 | 백업 실패·폴더 교체 시 변환 계속 + task·경고 명시 | `execute_task` (경고 추가 수정) | 통과 |
| P0-D1/D2 | planner vs MCP preview 분류·출력명 일치 (PDF/HWP) | `build_tasks` vs `preview_conversion` | 통과 |
| P1-V1~11 | validator 서명·크기·확장자·미완료 거부 단위 11건 | `validate_job_outputs` 스텁 레코드 | 통과 |
| P1-C1~7 | MCP 설정 fail-closed·TOML·환경변수·job_dir 7건 | `load_config` | 통과 |
| P1-D1/D2 | exit 0 + 성공 report + DOCX 내용 PDF → 검증 실패 / ZIP 교정 후 통과 | 실 subprocess fake CLI → validator | 통과 |
| CLI-CP949 | cp949 콘솔 ⚠️ 출력 충돌 → 종료코드 1 (실측 결함) | `_run_cli_conversion` | 수정 + 회귀 테스트 통과, 실 콘솔 재실행 exit 0 |

## 출력 증거 (실 COM)

합성 입력 2종(`sample_ko.hwp` 79,872B / `sample_ko.hwpx` 77,775B, 90줄 동일 내용).

| 입력 → 형식 | 결과 | 독립 오라클 |
|---|---|---|
| HWPX → PDF/DOCX/PNG/HTML/TXT | 5/5 성공 (25.7s) | PDF 3p·첫/끝줄/제목 추출 일치, PNG 3장 디코드(992×1403 RGB), DOCX `document.xml` 첫/끝줄 일치, HTML(cp949) 일치, TXT(cp949) 일치 |
| HWP → PDF/DOCX | 2/2 성공 | PDF 3p·내용 일치, DOCX 내용 일치, `saveas_3` 폴백 사용 기록 |
| 폴더 → PDF (실 CLI, UTF-8) | exit 0, 2/2 성공 | 동일 stem `sample_ko (1).pdf` 자동 개명 + report JSON |
| 폴더 → PDF (실 CLI, cp949 콘솔) | 수정 전 exit 1 충돌 / 수정 후 exit 0 | 회귀 테스트 `test_cli_survives_non_utf8_console_output` |
| 폴더 → DOCX (GUI 공용 Worker) | success 2/0/0/0 | DOCX 2건 내용 일치, `sample_ko (1).docx` 개명 |
| MCP preview→submit→result→validate (실 CLI) | succeeded 2/0, 검증 2/2 | `sample_ko.pdf` + `sample_ko (1).pdf` 내용 일치 |

- 입력 원본 SHA-256: 변환 전후 불변 확인 (2종 모두).
- 증거: `docs/quality/com_evidence/` (`smoke_hwpx.json`, `smoke_hwp.json`, `cli_report.json`, `fixture_manifest.json`).

## COM 동작·오류·MCP·Mutation·안전

- COM 동작: 앱 소유 PID만 생성·추적·종료 (실측 PID 27324·25836·22912 등, 종료 후 잔류 0건 확인). `HWPFrame.HwpObject`, `saveas_3` 폴백 실측.
- 호환 문서 팝업: 소유 PID 한정 자동 계속 실측 3회 (`HwpDialogAutoResponder`, pids=[소유]). raw-COM 픽스처 경로의 보안 팝업 1회는 사용자가 수동 '모두 허용'으로 해소 (제품 경로 아님, 보고서에 분리 기록).
- SaveAs 2-param 실패는 한글 2022 실측 특성 (팝업 없는 상태에서도 재현, 3-param 정상). 제품의 2→3-param 폴백 정당.
- 오류/타임아웃: MCP timeout·BUSY·REPORT_MISSING 기존 테스트 통과. 강제 종료는 `kill_owned_processes` 소유 한정 (테스트 `test_force_terminate_uses_owned_processes_only`).
- MCP: stdio 핸드셰이크 + preview/submit/result/validate 실측 (상용 CLI 래퍼). `overwrite=false`·명시 출력 폴더·승인 토큰·멱등 계약 유지.
- Mutation 8건 (전부 목표 테스트 적색 → 원복 녹색, 잔류 변경 0건): M1 원본보호 무력화 / M2 백업 접두사 정리 / M3 실행시점 재할당 생략 / M4 산출물 검증 생략 / M5 승인 우회 / M6 누락 report 성공 둔갑 / M7 서명 항상통과 / M8 PLAN_STALE 생략.
- 사용자 PC 안전: 실사용 문서·열린 한글·레지스트리(보안모듈 HKCU 사전준비 제외)·업데이트 EXE 변경 없음. 보안모듈 DLL은 제품 정상 동작 (`%LOCALAPPDATA%/HwpMate/security/`). QA 스크래치 삭제됨.

## 수행 불가·미검증·잔여 리스크

- 500건+ 실 COM 스트레스, 관리자 권한·frozen EXE 교체·실 레지스트리 등록 시험: 미실행 (`release_vm` 마커로 분리, 별도 승인 필요).
- 장시간 폴더 스캔 차단 E2E: 미실행 (코드에 스캔 캐시 완화 존재, 기존 단위 수준만 확인).
- `converter.py` COM 분기·보안모듈·GUI 컨트롤러 저커버리지 영역: 갭 인벤토리 참조.
- CI 실행 URL: 없음 (로컬 단독 실행). 재현 명령은 각 증거 JSON과 본 보고서 TEST-ID 참조.
