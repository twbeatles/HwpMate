# Project Audit

**감사 일자:** 2026-07-15  
**후속 구현:** 2026-07-15 (본 문서 권장안 1~3단계 반영 완료, `update_history.md` 참고)  
**대상:** HwpMate (`hwptopdf-hwpx_v4.py` → `hwpmate/`)  
**방법:** README.md / Claude.md 정독 → CodeGraph MCP 구조·호출 관계 분석 → 필요 시 핵심 소스 확인 → `pytest` / `pyright` 보조 검증  
**범위:** 기능 구현 관점 감사 후, 권장 수정안 구현 완료

---

## 1. Executive Summary

HwpMate는 한컴오피스 한글 COM 자동화를 쓰는 **Windows 전용 일괄 변환 GUI**다. 2026-05~06 구간에서 변환 중 busy guard, 보조 산출물 충돌 회피, 원자 저장, 단일 인스턴스 잠금, Open 실패 정리 등이 이미 보강되어 있고, 현재 자동 검증 기준은 안정적이다.

| 항목 | 결과 |
|------|------|
| `python -m pytest -q` | **67 passed** (약 7.4s) |
| `python -m pyright .` | **0 errors / 0 warnings** |
| 전체 기능 위험도 | **Medium** (일반 경로 Low–Medium, COM/대량 폴더/강제 종료 경로 Medium–High) |

**핵심 잔여 리스크 (코드 근거 있음):**

1. **폴더 모드 변환 시작 시 GUI 스레드 동기 재스캔** — 대용량 폴더에서 UI 정지 가능  
2. **취소 중 진행 파일 상태 분류 오류** — 실패 COM 오류가 있으면 `취소됨` 대신 `실패`로 집계  
3. **워커·컨버터 이중 `CoInitialize`/`CoUninitialize`** — 아파트 참조 카운트 불균형 가능성  
4. **강제 종료 PID 재사용 위험** — 추적 PID에 `taskkill /F`만 수행, 재사용 검증 없음  
5. **프로세스 추적 실패 시 강제 종료 불가** — 이미 실행 중인 한글에 attach 되면 응답 없음 상태로 남을 수 있음  

이전 감사(2026-06-11)에서 닫힌 항목(중복 시작, 보조 산출물 충돌, Open False Clear, 설정 저장 실패 전파 등)은 현재 코드에서도 유지되고 있다. 남은 문제는 **COM 환경 의존성**과 **대량 입력·취소·강제 종료 경계**에 집중되어 있다.

---

## 2. Project Understanding

### 2.1 목적과 규칙 (README / Claude.md)

- **목적:** HWP/HWPX → PDF, HWPX, DOCX, ODT, HTML, RTF, TXT, PNG/JPG/BMP/GIF 일괄 변환  
- **스택:** PyQt6 UI, pywin32 COM, PyInstaller 배포 (`HWP변환기_v8.7.exe`, `uac_admin=True`)  
- **환경:** Windows 10/11, Python 3.10+, 한컴오피스 한글, 관리자 권한 사실상 필수  
- **유지보수 중심:** `hwpmate/` 패키지. `legacy/hwptopdf-hwpx v3.py`는 참고용  
- **깨지면 안 되는 로직:** SaveAs 2→3 인자 폴백, 워커 COM 초기화, 보안 모듈 등록, 네이티브 DnD, 백업 실패 비중단, 성공=산출물 갱신·비어 있지 않음, 동일 형식 건너뜀, 단일 인스턴스, 소유 PID만 강제 종료  

### 2.2 아키텍처 (CodeGraph + 패키지 구조)

```text
hwptopdf-hwpx_v4.py
  └─ hwpmate.bootstrap.main
       └─ hwpmate.app.main
            ├─ pywin32 / is_admin / SingleInstanceLock
            ├─ enable_drag_drop_for_admin (정책 허용 시)
            └─ MainWindow
                 ├─ AppearanceController      (테마·형식·busy UI)
                 ├─ FileSelectionController   (스캔·파일 목록)
                 ├─ ConversionController      (계획·워커·결과)
                 ├─ NativeDropController      (WM_DROPFILES)
                 └─ LifecycleController       (메뉴·트레이·종료·설정)
```

**CodeGraph blast radius 요약 (대표):**

| 심볼 | 주요 호출자 / 영향 |
|------|-------------------|
| `ConversionWorker` | `MainWindow` 생성·시그널, 테스트 `test_conversion_worker` |
| `start_conversion` | 단축키/버튼 → 중복 시작 차단·preflight·워커 start |
| `HWPConverter.convert_file` | 워커 재시도 루프, SaveAs 폴백·artifact snapshot |
| `TaskPlanner.build_tasks` / `resolve_output_conflicts` | 변환 계획·충돌 리네임, `artifact_policy` 공유 |
| `save_config` / `ConfigRepository` | 종료·테마·변환 시작 시 설정 영속화 |
| `SingleInstanceLock` | `app.main` 단일 실행 |

### 2.3 주요 실행 흐름

1. **기동:** `app.main` → pywin32 검사 → 관리자 아니면 종료 → DnD 정책 → `QLockFile` → `MainWindow`  
2. **입력:** 폴더 미리보기/`FileScanWorker` 또는 파일 추가 스캔 → `FileSelectionStore` 중복 제거  
3. **계획:** `TaskPlanner.build_tasks` → 동일 확장자 건너뜀 → `resolve_output_conflicts` (기본+보조 산출물)  
4. **사전 점검:** `PreflightDialog` (입력 존재/읽기, 출력 쓰기, ProgID)  
5. **변환:** `ConversionWorker.run` → 워커 COM 초기화 → `HWPConverter.initialize` → 파일별 백업·재시도·`convert_file`  
6. **성공 판정:** Open/SaveAs False 거부, 전후 artifact snapshot 비교, size>0, Clear 정리  
7. **결과:** `ConversionSummary` → 토스트/`ResultDialog` → TXT·CSV·JSON 원자 저장  
8. **종료/취소:** cancel 플래그 → (타임아웃 시) `kill_owned_processes` → 설정 저장  

### 2.4 문서 vs 구현 정합 (요약)

| 문서 주장 | 구현 상태 |
|-----------|-----------|
| SaveAs 2/3 인자 폴백 | 일치 (`hwp_converter.convert_file`) |
| 변환 중 입력/드롭/시작 차단 | 일치 (컨트롤러 busy guard) |
| 보조 산출물 성공/충돌 반영 | 일치 (`artifact_policy`) |
| 단일 인스턴스 | 일치 (`app_instance`) |
| 소유 PID만 강제 종료 | 일치 (단, 추적 실패 시 비활성) |
| 사용법/정보 문구 “기본 출력 파일만 성공 기준” | **불일치** — 코드는 보조 산출물도 성공 인정 |
| README 지원 형식·버전 8.7 | 일치 (`constants.VERSION`, `FORMAT_TYPES`) |

---

## 3. High-Risk Issues

> 이전 감사에서 **완료**로 닫힌 항목은 재발 여부를 코드로 확인했고, 아래는 **현재 잔여** 이슈다.  
> 스타일/취향 지적은 제외했다.

### 3.1 폴더 모드 변환 시작 시 UI 스레드 동기 재스캔

* **위치:** `TaskPlanner.build_tasks` (folder 분기) ← `ConversionController.collect_tasks` ← `start_conversion`  
* **문제:** 폴더 미리보기는 `FileScanWorker`로 비동기인데, 변환 시작 시 `build_tasks`가 **메인(UI) 스레드에서 `iter_supported_files`로 전체 트리를 다시 순회**한다.  
* **영향:** 대량 파일/깊은 트리는 사전 점검 다이얼로그 전까지 UI 정지(응답 없음) 가능. 미리보기 결과와 실제 실행 목록도 시점 차로 어긋날 수 있다.  
* **근거:**
  - `file_selection.start_folder_preview_scan` → 비동기 `FileScanWorker`
  - `conversion.collect_tasks` → `task_planner.build_tasks` → 폴더 모드에서 동기 `iter_supported_files`  
* **권장 수정 방향:**
  - 미리보기 스캔 결과를 캐시해 `build_tasks` 입력으로 재사용하거나
  - 변환 계획 수립을 워커/비동기 단계로 분리하고 preflight 직전 UI 로딩 표시  
* **우선순위:** **High** (대용량 폴더 실사용 시)

### 3.2 취소 요청 중 진행 파일의 상태 분류 오류

* **위치:** `ConversionWorker.run` (재시도 루프 직후 상태 결정)  
* **문제:** 취소 후 진행 중 파일이 실패하면 다음 조건만 `취소됨`으로 본다.

```text
cancel_requested and not success and error is None  → 취소됨
success                                               → 성공
else                                                  → 실패  (error 메시지 유지)
```

사용자가 취소한 뒤 COM이 오류 문자열을 반환하면 **의도상 취소인데 실패로 집계**된다.  
* **영향:** 결과 다이얼로그/CSV의 실패·취소 카운트 왜곡, 재시도 정책 해석 혼란.  
* **근거:** `conversion_worker.py` 상태 분기; 취소 시 대기 작업만 `취소됨`으로 일괄 마킹.  
* **권장 수정 방향:** `cancel_requested and not success`이면 `취소됨`으로 통일하고, COM 오류는 `error`/`detail`에 보조 기록.  
* **우선순위:** **Medium**

### 3.3 워커와 컨버터의 이중 COM 아파트 초기화/해제

* **위치:**
  - `ConversionWorker.run`: `CoInitialize` / finally `CoUninitialize`
  - `HWPConverter.initialize` / `cleanup`: 동일 API 재호출  
* **문제:** 같은 워커 스레드에서 COM을 두 번 초기화·해제한다. `initialize`의 `CoInitialize` 예외는 무시하고, `cleanup`은 항상 `CoUninitialize`를 시도한다.  
* **영향:** 아파트 참조 카운트 불균형 시 이후 파일 변환 또는 강제 종료 후 잔존 스레드에서 COM 호출 실패·hang 가능. (환경/pywin32 버전에 따라 재현 강도 다름)  
* **근거:** CodeGraph/`grep`으로 양쪽 `CoInitialize`/`CoUninitialize` 확인. Claude.md는 워커 쪽 호출 유지를 요구하므로 **제거 시 한쪽만 소유**하도록 역할을 분리해야 한다.  
* **권장 수정 방향:**
  - 워커만 apartment 소유, 컨버터는 Dispatch/Quit만 담당 **또는**
  - 컨버터가 apartment를 소유하고 워커는 중복 호출 금지  
  - 중첩 시 `CoInitialize` 성공 여부를 플래그로 추적해 대칭 해제  
* **우선순위:** **Medium** (간헐적 COM 이슈 원인 후보)

### 3.4 강제 종료 시 PID 재사용 검증 부재

* **위치:** `HWPConverter.kill_owned_processes`  
* **문제:** 초기화 시점 스냅샷 diff로 `owned_pids`를 저장한 뒤, 강제 종료 시 **PID만으로 `taskkill /F`** 한다. 프로세스가 종료된 뒤 OS가 PID를 재사용하면 다른 프로세스에 강제 종료가 갈 수 있다.  
* **영향:** 드물지만 파괴적. 관리자 권한 앱이라 피해 반경이 큼.  
* **근거:** `_snapshot_hwp_pids` → set 차집합 저장 → `taskkill /PID` (이미지명/생성 시각 재검증 없음).  
* **권장 수정 방향:** 종료 직전 `tasklist`/WMI로 이미지명이 `Hwp.exe` 등인지 재확인; 가능하면 생성 시각·명령줄 교차 검증.  
* **우선순위:** **Medium** (확률 낮음, 영향 큼)

### 3.5 한글 프로세스 추적 실패 시 강제 종료·응답 없음 복구 공백

* **위치:** `HWPConverter.initialize` (`owned_pids` 공백 시 경고), `ConversionWorker.can_force_terminate` / `force_terminate`, `ConversionController.request_worker_stop`  
* **문제:** 이미 떠 있는 한글에 COM attach 되면 새 PID가 없어 `has_owned_processes()==False`. 이 경우 취소 후 강제 종료 UI가 비활성/불가이고, COM hang 시 워커·앱 종료가 막힐 수 있다.  
* **영향:** 변환 “응답 없음” 체감, close_event가 워커 대기로 `event.ignore()` 반복 가능.  
* **근거:** `process_tracking_warning` 설정 및 `can_force_terminate`가 converter owned PID에만 의존. Claude.md 정책상 전역 `Hwp.exe` kill은 금지.  
* **권장 수정 방향:**
  - 추적 실패를 preflight/상태바에 더 눈에 띄게 표시
  - hang 시 “마지막 수단” 옵션을 분리 설계(명시 동의 + 범위 제한)하거나 타임아웃 후 사용자 안내 강화
  - 초기화 전 사용자 한글 창 종료 권고  
* **우선순위:** **High** (운영 장애 체감), 다만 정책 트레이드오프 존재

### 3.6 스캔 취소 대기 시간 과소 (`SCAN_CANCEL_WAIT_MS = 200`)

* **위치:** `constants.SCAN_CANCEL_WAIT_MS`, `FileSelectionController.cancel_active_scan`, `ConversionController.start_conversion`  
* **문제:** 활성 스캔 취소 후 200ms만 대기한다. 대용량 스캔이 끝나지 않으면 변환 시작이 “스캔 종료 대기”로 거절되거나, 종료 이벤트가 무시된다.  
* **영향:** UX 마찰, 종료/변환 시작 실패 체감. 데이터 손상보다는 사용성/상태 꼬임.  
* **근거:** `cancel_active_scan(wait_ms=SCAN_CANCEL_WAIT_MS)` 기본 200; 변환 시작 시 folder_preview 취소 실패 분기.  
* **권장 수정 방향:** 대기 시간 상향 또는 진행 중 폴링/재시도, 스캔 중 변환 시작 시 “스캔 취소 후 자동 재개” UX.  
* **우선순위:** **Medium**

### 3.7 (잔존 환경 리스크) 실제 COM/빌드 미자동검증

* **위치:** `tools/hwp_com_smoke.py`, `HWP_COM_SMOKE_TEST_CHECKLIST.md`, `hwp_converter.spec`  
* **문제:** 단위 테스트는 COM을 모킹한다. 실제 한글 버전별 SaveAs/이미지·HTML 보조 산출물, UAC DnD, PyInstaller 바이너리는 CI/로컬 자동 검증 밖이다.  
* **영향:** 회귀가 릴리스 후에야 드러날 수 있음.  
* **근거:** 테스트 스위트에 실 COM 통합 테스트 없음; README/Claude가 수동 스모크를 권장.  
* **권장 수정 방향:** 관리자 환경 주기 스모크 체크리스트 실행, 가능하면 형식별 smoke matrix 최소화 자동화.  
* **우선순위:** **Medium** (제품 리스크, 코드 버그 단정은 아님)

---

## 4. Potential Functional Gaps

확실하지 않은 항목은 **추정**으로 표시한다.

### 4.1 구현 근거가 있는 보완 지점

| 항목 | 설명 |
|------|------|
| 미리보기 vs 실행 목록 정책 불일치 | 폴더 미리보기는 `preview_allowed_extensions`로 동일 형식 제외, `build_tasks`는 전체 지원 확장자를 모아 동일 형식을 **건너뜀**으로 넣음. 카운트 불일치 가능. |
| 진행률 표시 | `progress_updated.emit(idx, ...)`가 **작업 시작 시점** 인덱스라 첫 파일 동안 0/N으로 보임. 완료 기준 갱신이 아님. |
| 정보/사용법 문구 | `LifecycleController.show_about` / `show_usage`가 “기본 출력 파일 존재·0바이트 초과”만 성공 기준으로 설명. 실제는 보조 산출물 허용. |
| 설정 경로 분산 | 설정: `~/.hwp_converter_config.json`, 로그: `~/.hwp_converter/logs` 또는 `%LOCALAPPDATA%/HwpMate`, 락: `%LOCALAPPDATA%/HwpMate`. 동작은 하나 문서/지원 관점에서 파편화. |
| 파일 선택 필터 | `browse_files`에 `모든 파일 (*.*)` 허용 → 비대상 파일은 스캔에서 조용히 제외될 수 있음 (피드백 약함). |
| 취소 중 성공 | 취소 요청 후 이미 성공한 `convert_file`은 성공 유지(합리적). 실패+취소는 3.2 이슈. |

### 4.2 추정 (요구사항 미확인)

| 항목 | 추정 내용 |
|------|-----------|
| 암호/배포용 문서 | `forceopen:true`와 메시지박스 억제로 암호 문서·매크로 문서는 실패 또는 조용한 오동작 가능. 전용 UX 없음. **추정** |
| 네트워크/동기화 폴더 | OneDrive 등에서 스냅샷 mtime/size 경쟁으로 성공 판정 오탐/미탐 가능. **추정** |
| 병렬 변환 | 단일 COM/단일 워커 설계. 성능 향상을 위한 다중 인스턴스는 정책상 의도적 미지원으로 보임. **추정** |
| 결과 재변환 필터 | 실패 항목만 재실행 UI 없음. CSV 기반 수동 재시도. **추정(기능 공백 후보)** |
| 다국어/접근성 | UI 문자열 한국어 고정, 스크린리더 등은 범위 밖. **추정** |

### 4.3 이전 감사 대비 이미 닫힌 항목 (재확인)

다음 항목은 현재 코드·테스트에 방어 로직이 있어 **잔여 High-Risk로 올리지 않음:**

- 변환 중 중복 시작 / 입력·드롭·단축키 차단  
- 이미지·HTML 보조 산출물 충돌 회피  
- Open False 시 Clear  
- 설정/결과 원자 저장 및 저장 실패 UI 전파  
- 단일 인스턴스 잠금  
- Windows 경로 예약명·잘못된 문자 검증  

---

## 5. Recommended Fix Plan

### 1단계 — 즉시 수정 (기능 신뢰도·운영 장애)

1. **취소 상태 분류 수정** (`ConversionWorker`): `cancel_requested and not success` → `취소됨`  
2. **폴더 모드 변환 계획 비동기화 또는 미리보기 캐시 재사용** (UI freeze 제거)  
3. **프로세스 추적 실패 가시성 강화** (preflight/상태바 + 종료 불가 시 명확 안내)  
4. **강제 종료 전 PID 이미지명 재검증**  

### 2단계 — 안정성 개선

1. COM apartment 소유권 단일화 (워커 vs 컨버터)  
2. `SCAN_CANCEL_WAIT_MS` 상향 또는 스캔 취소 재시도 UX  
3. 진행률을 “완료 개수” 기준으로 변경  
4. About/Usage/README 성공 판정 문구를 보조 산출물 정책에 맞게 동기화  
5. 설정·로그·락 경로를 문서에 표로 정리 (가능하면 디렉터리 정책 통일)

### 3단계 — 구조 개선

1. 폴더 미리보기 결과 → `PlannedConversion` 파이프라인 단일화  
2. 실패 항목만 재실행 워크플로 (결과 다이얼로그 연계) **추정 요구**  
3. COM 스모크를 형식 매트릭스/체크리스트와 릴리스 게이트에 고정  
4. `file_list`/`_file_set` 별칭을 제거하고 `file_store` API만 사용 (유지보수 실수 방지)  
5. 컨트롤러 단위에서 cancel/close/force-kill 통합 시나리오 테스트 확대  

---

## 6. Test Recommendations

### 6.1 단위/회귀 (우선 추가)

| 테스트 | 목적 |
|--------|------|
| `ConversionWorker`: 취소 중 `convert_file`이 `(False, "COM error")` 반환 → status=`취소됨` | 3.2 회귀 |
| `ConversionWorker`: 취소 중 성공 → `성공` 유지 | 경계 확인 |
| `TaskPlanner`/`ConversionController`: 폴더 모드 계획 수립이 UI 스레드에서 장시간 walk 하지 않음 (캐시/워커 mock) | 3.1 |
| `kill_owned_processes`: 종료 직전 이미지명 불일치 시 taskkill 미호출 | 3.4 |
| `can_force_terminate` False일 때 close/cancel UI 메시지 경로 | 3.5 |
| `cancel_active_scan` 타임아웃 시 start_conversion 거절 메시지 | 3.6 |
| About/상수 문서 문자열이 아닌 **정책 함수**와 성공 기준 일치 스냅샷 | 문서 드리프트 방지 |

### 6.2 이미 있는 테스트 (유지)

- busy guard, 네이티브 드롭 차단, 보조 산출물 충돌, Open False Clear  
- 설정 정규화/원자 저장, 결과 CSV/JSON 필드, 단일 인스턴스  
- TaskPlanner 건너뜀/상대 경로/타임스탬프 폴백  

### 6.3 수동/릴리스 (자동 대체 곤란)

체크리스트는 기존 `HWP_COM_SMOKE_TEST_CHECKLIST.md`를 따르되, 감사 관점에서 특히:

1. 관리자 권한 기동 + 두 번째 인스턴스 차단  
2. 폴더 모드 1만+ 파일(또는 깊은 트리)에서 변환 시작 시 UI 응답성  
3. HWP/HWPX → PDF, DOCX, HTML, PNG 각 1건 + 보조 산출물 생성 환경  
4. 변환 중 Esc 취소 → 집계(취소/실패) 확인  
5. 한글 미실행 vs 이미 실행 상태에서 강제 종료 가능 여부  
6. 덮어쓰기 OFF + 기존 HTML/이미지 보조 산출물 충돌 리네임  
7. 결과 CSV/JSON 감사 필드 및 원자 저장  
8. `pyinstaller hwp_converter.spec` 산출물 실행 + DnD  

### 6.4 현재 자동 검증 스냅샷

```text
pytest  : 67 passed
pyright : 0 errors, 0 warnings, 0 informations
```

---

## 부록 A. CodeGraph로 확인한 호출 경로 (요약)

```text
app.main
  → SingleInstanceLock.try_lock
  → MainWindow
       → ConversionController.start_conversion
            → TaskPlanner.build_tasks / resolve_output_conflicts
            → PreflightDialog
            → ConversionWorker.start
                 → HWPConverter.initialize / convert_file / cleanup
                 → task_completed(ConversionSummary)
            → ResultDialog / write_* 원자 저장

NativeDropController.on_native_files_dropped
  → (busy guard) folder_preview | FileSelectionController.add_files

LifecycleController.close_event
  → cancel scan / request_worker_stop / force_terminate / save_settings
```

## 부록 B. 감사 범위 밖 / 의도적 비범위

- 레거시 `legacy/hwptopdf-hwpx v3.py` 품질 (유지보수 비대상)  
- 코드 스타일, 네이밍 취향  
- 비Windows 이식  
- 라이선스/배포 채널 운영  

---

*이 문서는 코드를 수정하지 않고 작성한 기능 구현 감사 리포트다. 수정 착수 시 1단계 항목부터 PR 단위로 처리하는 것을 권장한다.*
