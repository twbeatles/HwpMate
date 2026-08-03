# Project Audit

**감사 일자:** 2026-08-03 (재감사)  
**후속 구현:** 2026-08-03 — 본 문서 1~3단계 권장안 반영 (`update_history.md` 「재감사 권장안 구현」)  
**감사 초점:** 직전 감사 권장안 반영 **이후** 전체 기능 구현 상태  
**대상:** HwpMate (`hwptopdf-hwpx_v4.py` → `hwpmate/`)  
**방법:** README.md / Claude.md 정독 → CodeGraph MCP 구조·호출 관계 분석 → 필요 시 핵심 소스 확인 → `pytest` / `pyright` 보조 검증  

| 검증 (권장안 구현 후) | 결과 |
|------|------|
| `python -m pytest -q` | **98 passed** |
| `python -m pyright .` | **0 errors / 0 warnings** |
| 전체 기능 위험도 (구현 전 재감사) | **Low–Medium** → 계획 잠금·연결 구간 정책 반영 후 잔여는 COM 환경 의존 위주 |

---

## 1. Executive Summary

HwpMate는 한컴오피스 한글 COM으로 HWP/HWPX를 일괄 변환하는 Windows 전용 PyQt6 앱이다. 직전 감사 이후 **보안 모듈 SHA-256**, **소유 PID 우선 전면화**, **모듈 성공 시 자동 클릭 생략**, **폴더 스캔 wait 후 캐시 확정**, **`HwpSecuritySession`**, **자동 허용 설정 토글**, **pyright 클린**이 반영되어 이전 High 이슈의 상당수가 닫혔다.

**잔여 리스크는 “대량/경계/COM 환경”에 집중**된다. 자동 검증(93 tests + pyright)은 건강하다. 다만 연결 직후 폴링 창구, 스캔 대기 중 `processEvents` 재진입, 폴더 캐시 콜드 경로·신선도, 문서 드리프트는 여전히 기능적으로 문제를 만들 수 있다.

**후속 구현으로 해소된 항목 (2026-08-03):**  
1~4번(계획 잠금, 연결 전 자동 클릭 금지, 콜드 경로 비동기 스캔, 캐시 만료·신선도)은 코드에 반영됨.  

**구조적으로 남는 한계:**  
5. **프로세스 추적 실패(선실행 한글 attach) 시 강제 종료 불가** — 전체 Hwp kill 금지는 Claude.md 정책  

이전 감사 High(캐시 race, 전역 전면화+자동 클릭, pyright 실패)는 **현재 코드 기준으로 해소 또는 완화**된 것으로 재확인했다.

---

## 2. Project Understanding

### 2.1 목적과 규칙 (README / Claude.md)

| 항목 | 내용 |
|------|------|
| 목적 | HWP/HWPX → PDF·문서·이미지 일괄 변환 |
| 스택 | PyQt6, pywin32 COM, PyInstaller (`HWP변환기_v8.7.exe`, `uac_admin=True`) |
| 환경 | Windows 10/11, Python 3.10+, 한컴 한글, 관리자 권한 사실상 필수 |
| 깨지면 안 되는 로직 | SaveAs 2→3 폴백, 워커 COM apartment, 보안 모듈 실패 가시화, 네이티브 DnD, 백업 실패 비중단, artifact 성공 판정, 동일 형식 건너뜀, 단일 인스턴스, **소유 PID만** 강제 종료 |

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

**CodeGraph 호출 흐름 (핵심):**

`start_conversion` → (선택) `wait_for_active_scan` → `collect_tasks` → `TaskPlanner.build_tasks` → preflight → `ConversionWorker` → `HWPConverter.initialize`(`ensure_hwp_security_module`) → `engine_status_updated` → `HwpSecuritySession.apply_engine_status` → 파일 루프 `convert_file` → `ConversionSummary`.

| 심볼 | 영향 / 테스트 |
|------|----------------|
| `start_conversion` / `collect_tasks` | MainWindow 단축키·버튼; 컨트롤러 테스트 |
| `build_tasks` | conversion 다수 호출; `test_task_planner` |
| `HwpSecuritySession` | state·폴링; `test_hwp_security_session` |
| `ensure_hwp_security_module` | converter initialize; 단위 테스트(+무결성) |
| `kill_owned_processes` | 강제 종료; converter 테스트 있음 / Protocol 경로 약함 |

### 2.3 주요 실행 흐름

1. 기동: 관리자·단일 인스턴스·DnD 정책 → MainWindow  
2. 입력: 폴더 미리보기 캐시 또는 파일 스캔  
3. 계획: 캐시 우선 / 콜드 시 UI 재스캔, 동일 형식 건너뜀, 출력 충돌 조정  
4. 사전 점검: Preflight (+ 인쇄·허용 안내)  
5. 변환: 보안 DLL 설치·해시·레지스트리 → RegisterModule → Open/SaveAs → artifact 판정  
6. UI 폴링: 소유 PID 우선 전면화, 모듈 실패 시에만 자동 클릭(쿨다운·상한·옵션)  
7. 결과: Summary / 토스트 / ResultDialog / 원자 저장 CSV·JSON·TXT  

### 2.4 문서 vs 구현 정합

| 문서 주장 | 상태 |
|-----------|------|
| 보안 DLL + SHA-256 + LOCALAPPDATA | 일치 |
| 모듈 성공 시 자동 클릭 생략·옵션 off 가능 | 일치 |
| Toolhelp PID, 소유 PID 강제 종료 | 일치 |
| SaveAs 폴백·보조 산출물 성공 판정 | 일치 |
| `pyright` / `pytest` 통과 | **일치** (재감사 시점) |
| Claude.md 구조에 `hwp_security_module` / `hwp_security_session` | **미기재** (문서 갭) |
| `PROJECT_STRUCTURE_ANALYSIS.md` (2026-06 스냅샷) | **구식** — 보안 세션·SHA-256·설정 키 미반영 |
| README 구조 트리에 security 리소스 상세 | 부분 일치 (본문 기능 설명은 최신) |

---

## 3. High-Risk Issues

> 직전 감사에서 **해소된 항목**은 재발 여부만 확인하고, 아래는 **현재 잔여** 이슈다.

### 3.1 폴더 스캔 대기 중 `processEvents` 재진입

* **위치:** `FileSelectionController.wait_for_active_scan`, `ConversionController.start_conversion`
* **문제:** 스캔 대기 루프가 100ms마다 `QApplication.processEvents()`를 호출한다. 이 시점에는 `is_converting`이 아직 `False`라 busy guard가 없다. 대기 중 사용자가 변환 시작을 다시 누르거나 폴더/옵션을 바꿀 수 있다.
* **영향:** 중첩 `start_conversion`, 대기 중 입력 변경으로 계획/캐시 불일치, UI 상태 꼬임 (드묾~중간 빈도, 대용량 스캔에서 노출 시간 김).
* **근거:** `wait_for_active_scan`의 `processEvents` 루프; `start_conversion`의 `is_conversion_active` 검사는 wait **이전**에만 수행; wait 중 `set_converting_state(True)` 없음.
* **권장 수정 방향:** wait 시작 시 “계획 중/스캔 대기” 잠금 플래그로 시작 버튼·입력·드롭 차단; 또는 모달 progress + 단일 진입 가드(`_planning_in_progress`).
* **우선순위:** **High**

### 3.2 연결 직후 폴링 창: 전체 HWP + 자동 클릭 가능

* **위치:** `HwpSecuritySession.target_pids` / `should_auto_accept`; `ConversionController._start_hwp_foreground_polling` → `_poll_hwp_foreground`
* **문제:** 폴링은 워커 `start` 직후 시작되고, `engine_status_updated`는 `initialize` 완료 후에야 온다. 그 사이 `owned_pids` 비움 → `target_pids() is None` → 전면화/자동 클릭이 **시스템 전체 HWP** best-effort. 또한 `security_module_registered is None`이면 자동 클릭이 **허용**된다(`True`일 때만 차단).
* **영향:** 한글 연결에 수 초 걸리는 환경에서 다른 한글 문서 포커스 간섭·오클릭 여지. 이전 “전 구간 전역 폴링”보다는 짧지만 연결 창은 남음.
* **근거:** `_start_hwp_foreground_polling`에서 `reset_runtime()` 후 즉시 폴; `should_auto_accept`는 `registered is True`만 거부; `target_pids` 빈 집합 시 `None`.
* **권장 수정 방향:** 상태가 오기 전 자동 클릭 금지(`registered is not False`면 스킵 또는 “unknown” 정책); 전면화만 허용; 또는 initialize 시작 시그널 후 폴링 시작.
* **우선순위:** **Medium–High**

### 3.3 폴더 캐시 없는 콜드 경로의 UI 스레드 재스캔

* **위치:** `collect_tasks` (`require_folder_cache=False`), `TaskPlanner.build_tasks` `folder_file_paths is None` 분기
* **문제:** 설정 복원으로 폴더만 채워지고 미리보기 스캔이 안 된 채 변환하면 캐시 없이 UI 스레드에서 `iter_supported_files` 전체 순회.
* **영향:** 대용량 폴더에서 변환 시작 직후 UI freeze (이전 감사 3.1의 잔존 형태).
* **근거:** `main_window_ui`가 `folder_path`만 복원하고 스캔 미시작; `build_tasks` 재스캔 로그 경로 유지.
* **권장 수정 방향:** 시작 시 캐시 없으면 비동기 스캔 트리거 후 변환; 또는 복원 시 `QTimer.singleShot`으로 미리보기 스캔; 콜드 재스캔을 백그라운드 워커로.
* **우선순위:** **Medium**

### 3.4 폴더 스캔 캐시 신선도(stale cache)

* **위치:** `folder_scan_ready` / `folder_scan_files`; `get_folder_scan_cache`
* **문제:** 캐시는 폴더 경로·include_sub만 키로 삼고, mtime/TTL/내용 해시가 없다. 스캔 후 파일이 추가·삭제되어도 변환 계획은 옛 목록을 사용한다.
* **영향:** 누락 변환 또는 삭제된 파일 실패 다수 (사용자 인지 어려움).
* **근거:** `get_folder_scan_cache` 키 비교만 존재; invalidate는 폴더 재선택·include 변경·스캔 취소/오류 시.
* **권장 수정 방향:** 변환 직전 경량 재검증(개수·샘플 mtime) 또는 “캐시 시각 + N분” 만료; preflight에 캐시 시각 표시.
* **우선순위:** **Medium**

### 3.5 선실행 한글 attach 시 강제 종료 불가 (구조적)

* **위치:** `HWPConverter.initialize` owned_pids 차집합; `kill_owned_processes`; `can_force_terminate`
* **문제:** 이미 Hwp가 떠 있으면 새 PID가 안 잡혀 강제 종료가 비활성. Toolhelp 전환·스냅샷 경고 보강 후에도 attach 시나리오는 동일.
* **영향:** COM hang 시 사용자가 앱 소유 프로세스만으로 탈출 불가.
* **근거:** `owned_pids = after - before`; 비면 경고 + `has_owned_processes` False.
* **권장 수정 방향:** 연결 전 한글 종료 권고 강화; 세션 창 핸들 기반 보조 추적(신중); 전체 Hwp kill 복귀는 Claude.md 금지.
* **우선순위:** **Medium**

### 3.6 `error_occurred` 시그널 미사용 (죽은 경로)

* **위치:** `ConversionWorker.error_occurred`; `ConversionController.on_error_occurred`
* **문제:** 워커 `run()`은 예외를 `task_completed(summary)`로만 보내고 `error_occurred.emit`을 호출하지 않는다. UI 핸들러는 연결만 되어 있다.
* **영향:** 치명 오류 UI 분기가 사실상 사문화. 기능 회귀는 약하지만, 오류 UX 기대와 불일치.
* **근거:** 워커 전역 grep 상 `error_occurred.emit` 없음; 연결은 `_begin_worker_ui`에 존재.
* **권장 수정 방향:** 미사용 시그널 제거 또는 초기화 실패 등에서 emit 후 요약과 역할 분리 명확화.
* **우선순위:** **Low**

### 3.7 문서 드리프트 (Claude / 구조 분석)

* **위치:** `Claude.md` §3 구조; `PROJECT_STRUCTURE_ANALYSIS.md`
* **문제:** 신규 `hwp_security_module` / `hwp_security_session`, `auto_accept_security_dialog`, SHA-256 정책이 Claude 구조 목록·구조 분석 문서에 반영되지 않음.
* **영향:** 에이전트/기여자 오판, 회귀 시 문서 체크리스트 실패.
* **근거:** Claude.md services 목록에 security 모듈 없음; 구조 분석 설정 키에 `auto_accept` 없음, 일자 2026-06.
* **권장 수정 방향:** Claude.md·PROJECT_STRUCTURE_ANALYSIS를 현재 패키지와 동기화.
* **우선순위:** **Low–Medium**

---

## 4. Potential Functional Gaps

| 항목 | 구분 | 설명 |
|------|------|------|
| 스캔 대기 중 busy lock | **갭 (코드)** | 3.1 |
| 엔진 상태 전 자동 클릭 정책 | **갭 (코드)** | 3.2 — `registered is None`을 “미확인”으로 취급 안 함 |
| 설정 복원 시 자동 폴더 스캔 | **갭 (코드)** | 3.3 |
| 캐시 TTL / 디스크 동기화 | **갭 (코드)** | 3.4 |
| 폴더 작업 수집 백그라운드화 | **잔여 구조 (추정 보완)** | 콜드 경로 완전 제거에 필요 |
| Authenticode DLL 서명 검증 | **추정 보완** | SHA-256은 있음, 서명 체인은 없음 |
| COM 실기 자동화 | **환경 갭** | smoke·체크리스트 수동 |
| `error_occurred` 시그널 | **갭 (코드)** | 3.6 |
| 변환 중 입력 파일 삭제 | **기존 처리** | 워커가 파일 없음 실패 처리 — 양호 |
| 단일 인스턴스·원자 설정 저장 | **양호** | 유지 |
| 모듈 성공 시 폴링 완화·옵션 토글 | **양호** | 직전 구현 유효 |

---

## 5. Recommended Fix Plan

### 1단계 — 즉시 (안전·재진입)

1. 스캔 대기/`collect_tasks` 구간에 **계획 중 잠금** (`planning` 플래그)으로 시작·입력·드롭 차단  
2. `security_module_registered is None`일 때 **자동 클릭 금지** (전면화만 또는 폴링 지연)  
3. `engine_status_updated` 수신 전 전역 HWP(`pids=None`) 자동 클릭 금지  

### 2단계 — 안정성

1. 폴더 경로 복원 시 자동 미리보기 스캔 또는 변환 시 캐시 없으면 비동기 스캔 후 시작  
2. 캐시 만료·변환 직전 파일 존재 샘플 검증  
3. 스냅샷/추적 실패 시 preflight·상태바 문구 강화  
4. Claude.md / PROJECT_STRUCTURE_ANALYSIS 동기화  

### 3단계 — 구조

1. 폴더 작업 목록 생성을 스캔 워커와 동일 백그라운드 경로로 통합  
2. `error_occurred` 정리 또는 의미 있는 치명 경로에 연결  
3. (선택) DLL Authenticode 검증  
4. 계획 잠금·연결 구간 정책 회귀 테스트 고정  

---

## 6. Test Recommendations

### 6.1 우선 추가

| 테스트 | 목적 |
|--------|------|
| `wait_for_active_scan` 중 두 번째 `start_conversion` 무시 | 3.1 재진입 |
| `security_module_registered is None` → `should_auto_accept` False (정책 변경 후) | 3.2 |
| 폴링 시작 직후 `target_pids is None`일 때 `try_accept` 미호출 | 3.2 |
| 캐시 없이 폴더 변환 시 동기 재스캔 대신 스캔 요청/에러 | 3.3 정책 고정 |
| 캐시 후 파일 삭제 시 preflight/변환 실패 메시지 | 3.4 |
| `error_occurred` emit 여부 또는 제거 후 import 정리 | 3.6 |

### 6.2 유지·보강

| 영역 | 현황 |
|------|------|
| 보안 세션·폴링 PID | `test_hwp_security_session`, 컨트롤러 테스트 존재 — 유지 |
| DLL 무결성 | `test_hwp_security_module` — 유지 |
| kill_owned live PID | `test_hwp_converter` — 유지 |
| 설정 `auto_accept_security_dialog` | load/save round-trip 테스트 추가 권장 |

### 6.3 수동 스모크 (체크리스트 연계)

1. 대용량 폴더 스캔 중 변환 연타 → 중복 워커 없음  
2. 한글 선실행 + 변환 → 추적 경고·강제 종료 제한 안내  
3. 연결 수 초 구간 다른 한글 문서 편집 → 포커스 간섭 정도  
4. 설정만 복원된 폴더 즉시 변환 → freeze 여부  
5. 보안 모듈 정상 PC에서 자동 클릭 옵션 on/off  
6. `pyinstaller` 빌드 + DLL 포함·해시 통과  

---

## 부록 A. 직전 감사 대비 상태

| 2026-08-03 1차 감사 이슈 | 재감사 상태 |
|--------------------------|-------------|
| 폴더 wait 후 캐시 race | **해소** (processEvents 드레인 + require_folder_cache) |
| try_accept 전역 top-level 전면화 | **완화** (보안 창 위주 + 소유 PID) — 연결 창 잔존(3.2) |
| 모듈 성공 후 전 구간 폴링 자동 클릭 | **해소** |
| pyright 6 errors | **해소** |
| DLL 해시 없음 | **해소** (SHA-256) |
| UI 동기 재스캔 | **부분** — wait 경로 해소, 콜드 경로 잔존(3.3) |
| attach 시 강제 종료 | **잔존** |

## 부록 B. 검증 명령 (재감사 시점)

```text
python -m pytest -q
# 93 passed in ~6.3s

python -m pyright .
# 0 errors, 0 warnings, 0 informations
```

## 부록 C. 잘 유지되는 안전장치

- SaveAs 2/3 폴백, Open False 시 Clear  
- 보조 산출물 성공/충돌 정책 (`artifact_policy`)  
- 단일 인스턴스, 변환 중 busy guard (대기 구간 제외)  
- 결과/설정 원자 저장  
- 소유 PID 재확인 후 taskkill + CREATE_NO_WINDOW  
- 보안 모듈 실패 경고를 Summary에 포함  

---

*이 문서는 기능 재감사 결과물이다. 코드 변경은 포함하지 않았다.*
