# Project Audit

**감사 일자:** 2026-08-04  
**권장안 구현:** 2026-08-04 — 본 문서 1~3단계 핵심 항목 반영 (`update_history.md` 「PROJECT_AUDIT 권장안 구현」)  
**대상:** HwpMate (`hwptopdf-hwpx_v4.py` → `hwpmate/`)  
**방법:** `README.md` / `Claude.md` 정독 → CodeGraph MCP 호출 관계·영향 범위 분석 → 핵심 소스 보조 확인 → `pytest` / `pyright` 검증  

| 검증 | 결과 |
|------|------|
| `python -m pytest -q` | **107 passed** (~6.6s) |
| `python -m pyright .` | **0 errors / 0 warnings** |
| 전체 기능 위험도 | **Low** (구조적 COM 한계·대량 계획 백그라운드화·Authenticode는 잔여/보류) |

---

## 1. Executive Summary

HwpMate는 한컴오피스 한글 COM으로 HWP/HWPX를 일괄 변환하는 Windows 전용 PyQt6 앱이다. 패키지 분리(`services` / `workers` / `ui/main_window_controllers`), SaveAs 2→3 폴백, 소유 PID 강제 종료, 보안 DLL SHA-256·레지스트리, `is_planning` 재진입 가드, 폴더 스캔 캐시 TTL/신선도 샘플, 결과·설정 원자 저장 등 **핵심 안전장치가 이미 구현·테스트되어 있다.**

직전 감사(2026-08-03)에서 High로 지적됐던 항목 중 상당수는 **현재 코드 기준으로 해소 또는 의도적으로 완화**된 상태다.

| 과거 이슈 | 현재 상태 |
|-----------|-----------|
| 스캔 대기 중 변환 재진입 | **완화** — `start_conversion` 초기에 `_set_planning(True)` + `apply_busy_ui` |
| `engine_status` 전 자동 클릭 | **해소** — `should_auto_accept`가 `engine_status_received`·`registered is False`만 허용 |
| 폴더 콜드 경로 UI 동기 재스캔 | **해소** — `_ensure_folder_scan_ready` 비동기 스캔 + 캐시 필수 |
| 캐시 만료·부재 파일 | **부분 해소** — TTL 300s + 샘플 신선도 (한계 잔존) |
| DLL SHA-256 | **해소** |
| pyright 실패 | **해소** |

**남은 리스크는 “경계 재진입·대량 UI 작업·COM 환경 의존”에 집중**된다. 자동 검증은 건강하나, 실제 한글 COM 스모크와 대용량 폴더 시나리오는 여전히 수동/환경 의존이다.

**핵심 잔여 문제 (요약):**

1. 계획 중(`is_planning`) 종료 경로가 잠기지 않아 `processEvents` 대기 중 창 파괴 가능  
2. 대량 폴더에서 `collect_tasks`·사전 점검 상세가 UI 스레드에서 전체 파일 I/O  
3. 선실행 한글 attach 시 강제 종료 불가 (정책상 의도, 사용자 탈출 수단 제한)  
4. 폴더 캐시가 “신규 추가 파일”을 감지하지 못함  
5. COM/통합 테스트·Spec Kit 문서 포인터 드리프트  

---

## 2. Project Understanding

### 2.1 목적과 규칙 (README / Claude.md)

| 항목 | 내용 |
|------|------|
| 목적 | HWP/HWPX → PDF·문서·이미지 일괄 변환 |
| 스택 | PyQt6, pywin32 COM, PyInstaller (`HWP변환기_v9.0.exe`, `uac_admin=True`) |
| 환경 | Windows 10/11, Python 3.10+, 한컴 한글 2018+, 관리자 권한 사실상 필수 |
| 깨지면 안 되는 로직 | SaveAs 2→3 폴백, 워커 COM apartment, 보안 모듈 실패 가시화, 네이티브 DnD, 백업 실패 비중단, artifact 성공 판정, 동일 형식 건너뜀, 단일 인스턴스, **소유 PID만** 강제 종료, `is_planning`으로 스캔/계획 중 재진입 차단 |

### 2.2 아키텍처 (CodeGraph + 패키지)

```text
hwptopdf-hwpx_v4.py
  └─ hwpmate.app.main
       ├─ pywin32 · is_admin · SingleInstanceLock
       ├─ enable_drag_drop_for_admin (정책)
       └─ MainWindow
            ├─ FileSelectionController  (스캔·캐시·wait+processEvents)
            ├─ ConversionController     (계획·워커·HwpSecuritySession 폴링)
            ├─ Appearance / NativeDrop / Lifecycle
            └─ ToastManager
```

**CodeGraph 기준 핵심 호출 흐름:**

```text
start_conversion
  → _set_planning(True)
  → wait_for_active_scan / _ensure_folder_scan_ready
  → collect_tasks → TaskPlanner.build_tasks
  → PreflightDialog
  → ConversionWorker.run
       → CoInitialize
       → HWPConverter.initialize (ensure_hwp_security_module, RegisterModule, owned_pids)
       → engine_status_updated → HwpSecuritySession.apply_engine_status
       → convert_file (Open / SaveAs 2→3 / artifact snapshot)
  → ConversionSummary → ResultDialog
```

| 심볼 | 역할 | 테스트 |
|------|------|--------|
| `start_conversion` / `collect_tasks` | 계획·재진입 가드 | `test_main_window_controllers` |
| `TaskPlanner.build_tasks` | 건너뜀·출력 경로 | `test_task_planner` |
| `HwpSecuritySession` | 전면화/자동 클릭 정책 | `test_hwp_security_session` |
| `ensure_hwp_security_module` | DLL·SHA-256·HKCU | `test_hwp_security_module` |
| `kill_owned_processes` | 소유 PID만 taskkill | `test_hwp_converter` |
| `FileScanWorker` | 비동기 스캔 | 컨트롤러 단위 위주 (워커 자체 약함) |

### 2.3 주요 실행 흐름

1. **기동:** 관리자·단일 인스턴스·DnD 정책 → MainWindow → (지연) 네이티브 드롭·복원 폴더 미리보기 스캔  
2. **입력:** 폴더 미리보기 캐시 또는 파일 모드 비동기 스캔  
3. **계획:** `is_planning` 잠금 → 캐시 확보 → 동일 형식 건너뜀 → 출력 충돌 조정  
4. **사전 점검:** Preflight (입력 존재/읽기·출력 쓰기·ProgID·경고)  
5. **변환:** 보안 DLL 설치·해시·레지스트리 → RegisterModule → Open/SaveAs → artifact 판정·재시도·백업  
6. **UI 폴링:** 소유 PID 우선 전면화, 모듈 **실패** 시에만 「모두 허용」 자동 클릭 (쿨다운·상한·옵션)  
7. **결과:** Summary / 토스트 / ResultDialog / 원자 저장 CSV·JSON·TXT / 실패 재변환  

### 2.4 문서 vs 구현 정합

| 문서 주장 | 상태 |
|-----------|------|
| 보안 DLL + SHA-256 + LOCALAPPDATA | **일치** |
| 모듈 성공 시 자동 클릭 생략·옵션 off | **일치** (`should_auto_accept`) |
| Toolhelp PID, 소유 PID 강제 종료 | **일치** |
| SaveAs 폴백·보조 산출물 성공 판정 | **일치** |
| `is_planning`으로 스캔/계획 중 재진입 차단 | **부분 일치** — 입력/시작/드롭은 차단, **종료(close/tray)는 미차단** |
| 폴더 캐시 만료·샘플 검증 | **일치** (한계는 §3.5) |
| `pyright` / `pytest` 통과 | **일치** (본 감사 시점 98 passed) |
| Claude.md Spec Kit `specs/001-...` | **불일치** — 워크스페이스에 `specs/` 디렉터리 없음 |
| Claude.md services 목록 (security 모듈) | **일치** (현재 Claude.md 반영됨) |
| `PROJECT_STRUCTURE_ANALYSIS.md` | 대체로 최신(2026-08-03 보강). 본 감사 잔여 이슈는 미반영 |

---

## 3. High-Risk Issues

> 실제 코드 근거가 있는 **현재 잔여** 이슈만 기술한다.  
> 해소된 과거 이슈는 부록 A 참고.

### 3.1 계획 중 종료(`close` / 트레이)와 `processEvents` 재진입

* **위치:** `LifecycleController.close_event` (`lifecycle.py`); `FileSelectionController.wait_for_active_scan`; `ConversionController.start_conversion`
* **문제:** 변환 시작 시 `_set_planning(True)`로 시작·입력·드롭은 막지만, `close_event`는 `is_planning`을 보지 않는다. 스캔 대기 루프는 100ms마다 `QApplication.processEvents()`를 호출하므로, 대기 중 트레이 종료·Alt+F4가 들어오면 창이 닫히고 `start_conversion` 스택이 파괴된 위젯에 접근할 수 있다.
* **영향:** 종료 중 예외, 스캔 워커 정리 불완전, 설정 저장 타이밍 꼬임, 드묾~중간 빈도(대용량 스캔 대기 시간일수록 노출↑).
* **근거:**
  * `start_conversion`은 wait **전**에만 planning 잠금 후 `wait_for_active_scan` 진입
  * `wait_for_active_scan` 내부 `processEvents` 루프
  * `close_event`는 스캔 취소·`is_converting`만 처리하고 `is_planning` 분기 없음
* **권장 수정 방향:**
  1. `close_event`에서 `is_planning`이면 안내 후 ignore, 또는 계획 취소 플래그 후 안전 종료  
  2. wait 루프에서 `window` 유효성/`close_after_plan` 플래그 확인  
  3. 트레이 종료도 동일 가드
* **우선순위:** **High**

### 3.2 대량 폴더에서 UI 스레드 전체 작업·사전 점검 I/O

* **위치:** `ConversionController.collect_tasks` → `TaskPlanner.build_tasks`; `PreflightDialog._build_detail_text` / `_blocking_errors` / `_is_readable`
* **문제:** 폴더 스캔 자체는 백그라운드이나, 변환 직전 계획 수립과 사전 점검 상세 생성이 **UI 스레드**에서 전 파일 목록을 순회한다. Preflight는 실행 대상마다 `exists`·`open("rb").read(1)`로 읽기 가능 여부까지 검사한다.
* **영향:** 수천~수만 파일 배치에서 “변환 시작” 직후 UI freeze, 사전 점검 창 표시 지연, 디스크 부하. 사용자는 앱이 멈춘 것으로 느낄 수 있다.
* **근거:** `build_tasks` 캐시 경로가 메인 스레드에서 전체 `ConversionTask` 생성; Preflight `_is_readable`이 파일 단위 I/O.
* **권장 수정 방향:**
  1. 계획 수립/차단 오류 검사를 워커 또는 배치 샘플로 이동  
  2. Preflight 상세는 상위 N개 + “외 M개” 요약, 전체 목록은 접기/파일 저장  
  3. 진행 표시(모달 progress)로 대기 구간 가시화
* **우선순위:** **High** (대량 사용 시나리오 기준) / 소량 사용 시 Medium

### 3.3 연결 직후 `owned_pids` 부재 구간 전역 HWP 전면화

* **위치:** `HwpSecuritySession.target_pids`; `ConversionController._start_hwp_foreground_polling` / `_poll_hwp_foreground`; `bring_hwp_windows_to_foreground`
* **문제:** 폴링은 워커 start 직후 시작되고, `engine_status_updated`는 `initialize` 완료 후다. 그 사이 `target_pids() is None` → 시스템 전체 Hwp/HwpCtrl PID best-effort 전면화. **자동 클릭은 `engine_status_received` 전에는 금지**되어 이전 대비 안전하나, 다른 한글 세션의 보안/대화 창 포커스 간섭 여지는 남는다.
* **영향:** 연결에 수 초 걸리는 환경에서 사용자가 편집 중이던 다른 한글 창의 대화상자가 앞으로 올 수 있음. 오클릭 위험은 자동 클릭 정책으로 완화됨.
* **근거:** `reset_runtime()` 후 즉시 `_poll_hwp_foreground`; `target_pids` 빈 집합 시 `None`; `bring_hwp_windows_to_foreground(None)` → `_snapshot_hwp_pids()` 전체.
* **권장 수정 방향:** `engine_status_received` 전 전면화도 생략하거나, 연결 전용 상태 메시지 후 폴링 시작; 또는 initialize 시작 시그널을 워커에서 먼저 발행.
* **우선순위:** **Medium**

### 3.4 선실행 한글 attach 시 강제 종료 불가 (구조적)

* **위치:** `HWPConverter.initialize` (`owned_pids = after - before`); `kill_owned_processes`; `ConversionWorker.can_force_terminate`
* **문제:** 이미 Hwp가 떠 있으면 새 PID 차집합이 비어 강제 종료가 비활성. Claude.md 정책상 전체 `Hwp.exe` kill 금지와 일치하는 **의도된 안전장치**이나, COM hang 시 사용자 탈출 수단이 약하다.
* **영향:** 응답 없는 변환 세션에서 앱 취소/강제 종료 버튼이 실질적으로 무력화될 수 있음. 경고 문구·preflight 안내는 있음.
* **근거:** `owned_pids` 차집합; 비면 `process_tracking_warning` + `has_owned_processes` False.
* **권장 수정 방향:** 연결 전 “다른 한글 종료” 권고 강화(이미 일부 존재); 세션 창 핸들 기반 보조 추적(신중); 전체 kill 복귀 금지 유지.
* **우선순위:** **Medium**

### 3.5 폴더 스캔 캐시 신선도 — 신규 파일·후반부 삭제 미검출

* **위치:** `FileSelectionController.get_folder_scan_cache`; `validate_folder_scan_cache_freshness`
* **문제:** TTL(300초)과 **앞쪽 샘플 N개(24)** 존재 여부만 본다. (1) 스캔 이후 추가된 파일은 누락 변환, (2) 샘플 구간은 존재하고 뒤쪽만 삭제된 경우 계획에 잔존 → 변환 시점 “파일 없음” 실패.
* **영향:** 스캔 후 폴더 내용이 바뀐 작업 환경에서 결과 불일치. 완전 실패보다는 누락/부분 실패.
* **근거:** 샘플 `paths[:limit]`; 추가 파일 비교 로직 없음; TTL만 시간 기반.
* **권장 수정 방향:** 변환 직전 경량 재스캔 옵션; 샘플 랜덤/양 끝; 디렉터리 mtime 비교; preflight에 캐시 시각 표시.
* **우선순위:** **Medium**

### 3.6 폴더 모드 `relative_to` 예외 미포착

* **위치:** `TaskPlanner.build_tasks` (`input_file.relative_to(folder)`)
* **문제:** 캐시 경로가 선택 폴더 트리 밖으로 해석되면(`ValueError`) 계획 전체가 예외로 중단된다. 정규화는 대체로 일치하지만 교차 볼륨·특수 경로·수동 상태 조작 등 경계에서 가능.
* **영향:** 사용자에게 일반 예외 메시지; 해당 배치 변환 불가.
* **근거:** `relative_to` 호출부에 try/except 없음; `start_conversion`은 `Exception`으로 critical 박스 표시.
* **권장 수정 방향:** `relative_to` 실패 시 해당 파일 skip+경고 또는 flat 출력 폴백; 계획 단계에서 명시적 `ValueError` 메시지.
* **우선순위:** **Low–Medium**

### 3.7 `Open` 결과 판정이 `is False` 한정

* **위치:** `HWPConverter.convert_file`
* **문제:** `if open_result is False`만 실패로 본다. COM 버전에 따라 실패가 `0`/`None` 등으로 오면 identity 비교에 걸리지 않을 수 있다 (`0 is False` → False).
* **영향:** 열기 실패를 뒤쪽 SaveAs/artifact 단계에서야 실패로 집계; 오류 메시지 품질 저하. **추정 빈도: 환경 의존.**
* **근거:** `open_result is False` 분기; SaveAs는 `is False`와 예외 모두 처리.
* **권장 수정 방향:** 실패를 `open_result in (False, 0)` 또는 명시적 성공 집합으로 정규화; 스모크로 버전별 반환값 기록.
* **우선순위:** **Low–Medium** (추정 포함)

### 3.8 변환 중 `ConversionTask` 공유 뮤테이션 (스레드)

* **위치:** `ConversionWorker.run`이 worker 스레드에서 `task.status` 등 직접 변경; UI는 동일 객체 참조
* **문제:** 락 없이 공유 dataclass를 워커가 갱신한다. 현재 UI는 진행을 시그널 인자로만 표시하고 요약은 `task_completed` 이후 읽으므로 **실질 충돌 확률은 낮지만**, 향후 UI가 중간 `plan.tasks`를 읽으면 레이스가 된다.
* **영향:** 이론상 상태 표시 불일치; 현재 코드 경로는 대체로 안전.
* **근거:** 워커 루프 내 `task.status = ...`; Qt 시그널로 summary 전달.
* **권장 수정 방향:** 완료 시 불변 스냅샷/복사본 emit; 또는 UI 읽기 지점 문서화.
* **우선순위:** **Low**

### 3.9 Spec Kit / 감사 문서 드리프트

* **위치:** `Claude.md` SPECKIT 블록; 루트 `specs/`; 기존 `PROJECT_AUDIT.md` 서술
* **문제:** Claude.md는 `specs/001-hwp-mate-reliability-ux`와 tasks 상태를 가리키나 저장소에 `specs/`가 없다. 직전 감사 문서의 “잔여 High” 중 일부는 이미 코드로 닫혀 문서가 현실을 앞지르거나 뒤처질 수 있다.
* **영향:** 에이전트/기여자 오판, 중복 구현 또는 이미 고친 이슈 재수정.
* **근거:** `specs` 디렉터리 부재; 본 감사 시점 코드와 구 감사 3.1~3.4 서술 불일치.
* **권장 수정 방향:** Spec 포인터 정리 또는 실제 specs 복원; 감사 문서 일자·해소 표 유지.
* **우선순위:** **Low**

---

## 4. Potential Functional Gaps

| 항목 | 구분 | 설명 |
|------|------|------|
| 계획 중 close/tray 가드 | **갭 (코드)** | §3.1 |
| 대량 Preflight/계획 UI 부하 | **갭 (코드)** | §3.2 |
| 엔진 상태 전 전역 전면화 | **갭 (코드)** | §3.3 — 자동 클릭은 해소 |
| attach 시 강제 종료 | **구조적 한계** | §3.4 — Claude 정책과 트레이드오프 |
| 캐시 신규 파일 미반영 | **갭 (코드)** | §3.5 |
| `relative_to` 방어 | **갭 (코드)** | §3.6 |
| Open 반환값 정규화 | **추정 보완** | §3.7 |
| 긴 경로(MAX_PATH)/네트워크 경로 | **추정 보완** | HWP COM·Windows API 한도 문서화·검증 부족 |
| Authenticode DLL 서명 | **추정 보완** | SHA-256 번들 해시만 존재, 코드서명 체인 없음 |
| COM 실기 자동화 CI | **환경 갭** | `tools/hwp_com_smoke.py`·체크리스트 수동 |
| `FileScanWorker` 전용 단위 테스트 | **테스트 갭** | 취소·에러 시 `scan_finished` 미발행 경로 등 |
| `retry_failed_tasks` planning 잠금 | **낮은 갭** | 결과 다이얼로그 accept 후 호출; 모달 preflight로 대부분 보호 |
| 테마 토글 busy 미잠금 | **낮은 갭** | 변환 중 설정 저장 경쟁 가능 (기능 치명도 낮음) |
| 동일 형식 건너뜀·원자 저장·단일 인스턴스 | **양호** | 유지 |
| 모듈 성공 시 자동 클릭 생략·옵션 | **양호** | 유지 |
| 백업 실패 비중단 | **양호** | Claude 정책과 일치 |

---

## 5. Recommended Fix Plan

### 1단계 — 즉시 (안전·재진입)

1. **`is_planning` 중 종료 차단/순차 종료** — `close_event`·트레이 quit가 계획 대기 중 창을 바로 파괴하지 않도록  
2. **대량 배치 Preflight 부하 완화** — 상세 전체 `read(1)` 제거 또는 샘플링·비동기화; 목록 요약  
3. (선택) 계획 wait 루프에 cancel/closing 플래그 검사  

### 2단계 — 안정성

1. `engine_status_received` 전 전면화 폴링 지연 또는 소유 PID 확보 후에만 시작  
2. 캐시 신선도: 랜덤 샘플·신규 파일 감지 또는 “변환 전 빠른 재스캔” 옵션  
3. `relative_to` 실패 시 사용자 친화 메시지 / 파일 단위 스킵  
4. Open 반환값 실패 집합 정규화 + 스모크 체크리스트 기록  
5. Claude.md Spec Kit 포인터와 실제 트리 동기화  

### 3단계 — 구조

1. 폴더 작업 계획(`build_tasks`)을 스캔 워커와 동일 백그라운드 파이프라인으로 통합  
2. 워커→UI 결과 객체 불변 스냅샷화  
3. (선택) DLL Authenticode 검증  
4. 대용량·close-during-planning·캐시 드리프트 회귀 테스트 고정  

---

## 6. Test Recommendations

### 6.1 우선 추가

| 테스트 | 목적 |
|--------|------|
| `is_planning=True` 상태에서 `close_event` → ignore 또는 안전 종료 (정책 확정 후) | §3.1 |
| `wait_for_active_scan` 중 close/tray 시뮬레이션 | §3.1 |
| 대량 fake paths로 Preflight 생성 시간/호출 횟수 상한 (mock `_is_readable`) | §3.2 |
| `engine_status` 수신 전 `_poll`에서 `bring_hwp(..., None)` 호출 여부 (정책 변경 후) | §3.3 |
| 캐시 샘플 앞쪽만 존재·뒤쪽 삭제 / 신규 파일 추가 시 기대 동작 | §3.5 |
| `build_tasks`에서 folder 밖 path → 명확한 ValueError 메시지 | §3.6 |
| `Open`이 `0`/`False`/`True`일 때 convert_file 분기 | §3.7 |

### 6.2 유지·보강

| 영역 | 현황 |
|------|------|
| planning 중 start 차단 | `test_start_conversion_blocked_while_planning` — 유지 |
| 보안 세션 자동 클릭 조건 | `test_hwp_security_session` — 유지 |
| 캐시 신선도 기본 | `test_folder_cache_freshness_detects_missing_files` — 샘플 경계 케이스 보강 |
| DLL 무결성 | `test_hwp_security_module` — 유지 |
| kill_owned live PID | `test_hwp_converter` — 유지 |
| 네이티브 드롭 busy | `test_native_drop_ignores_paths_while_converting` — planning 중 드롭 케이스 추가 권장 |

### 6.3 수동 스모크 (체크리스트 연계)

1. 대용량 폴더 스캔 중 변환 시작 → 계획 잠금 UI, 종료 시도 시 동작 확인  
2. 1만 파일급 폴더 변환 시작 → Preflight 표시까지 freeze 여부  
3. 한글 선실행 + 변환 → 추적 경고·강제 종료 제한  
4. 연결 수 초 구간 다른 한글 문서 편집 → 포커스 간섭 정도  
5. 스캔 후 파일 추가/삭제 → 변환 목록 일치성  
6. 보안 모듈 정상/실패 PC에서 자동 클릭 옵션 on/off  
7. `python tools/hwp_com_smoke.py ...` 관리자 권한  
8. `pyinstaller hwp_converter.spec` + DLL 포함·해시 통과  

---

## 부록 A. 직전 감사 대비 상태

| 2026-08-03 재감사 이슈 | 2026-08-04 상태 |
|------------------------|-----------------|
| 3.1 processEvents 재진입 (start 중복) | **완화** — `is_planning`+busy UI; **종료 경로 잔존** → 본 문서 3.1 |
| 3.2 연결 창 전역 자동 클릭 | **해소** (전면화만 잔존 → 3.3) |
| 3.3 콜드 UI 재스캔 | **해소** (`_ensure_folder_scan_ready`) |
| 3.4 캐시 stale | **부분** — TTL·샘플 있음, 신규/후반 삭제 한계 → 3.5 |
| 3.5 attach 강제 종료 | **잔존** → 3.4 |
| 3.6 `error_occurred` 죽은 시그널 | **해소로 보임** — 현 코드에 시그널/핸들러 잔존 흔적 없음 |
| 3.7 문서 드리프트 | **부분** — Claude services 반영; Spec Kit 포인터·감사 문서 갱신 필요 |

---

## 부록 B. 검증 명령 (권장안 구현 후)

```text
python -m pytest -q
# 107 passed in ~6.6s

python -m pyright .
# 0 errors, 0 warnings, 0 informations
```

---

## 부록 C. 잘 유지되는 안전장치

- SaveAs 2/3 폴백, Open 실패 시 best-effort `Clear`  
- 보조 산출물 성공/충돌 정책 (`artifact_policy`)  
- 단일 인스턴스 (`QLockFile`), 변환·계획 중 입력/시작/드롭 busy guard  
- 결과/설정 원자 저장 (임시 파일 후 replace)  
- 소유 PID 재확인 후 `taskkill` + `CREATE_NO_WINDOW`  
- 보안 모듈 SHA-256 + HKCU 등록 + 실패 시 Summary 경고  
- 모듈 성공 시 자동 클릭 생략, 실패 시에만 쿨다운·상한 적용  
- 폴더 스캔 시 하위 `backup/` 기본 제외  
- 관리자 권한 강제, 네이티브 DnD 정책(IDLE 비활성 등)  

---

## 부록 D. CodeGraph 활용 메모

- 인덱스: 약 53 files  
- `start_conversion` → `collect_tasks` → 캐시/`build_tasks` 경로 확인  
- `ensure_hwp_security_module` → `verify_dll_integrity` → `sha256_file` 경로 확인  
- Blast radius 참고: `convert_file` / `FileScanWorker` 는 단위 테스트 표기가 약할 수 있으나 상위 컨트롤러·`test_hwp_converter`로 간접 커버  

---

*이 문서는 기능 구현 관점 감사 결과물이다. 코드 변경은 포함하지 않았다.*
