# Project Audit

**감사 일자:** 2026-08-04  
**감사 초점:** 인쇄 설정 제어 기능(`hwp_print_settings` + `convert_file` PDF 경로) 및 관련 영향 범위  
**대상:** HwpMate (`hwptopdf-hwpx_v4.py` → `hwpmate/`)  
**방법:** `README.md` / `Claude.md` 정독 → CodeGraph MCP 호출 관계·영향 범위 → 핵심 소스 보조 확인 → `pytest` / `pyright`  

| 검증 | 결과 |
|------|------|
| `python -m pytest -q` | **132 passed** (개선 반영 후) |
| `python -m pyright .` | **0 errors / 0 warnings** |
| 전체 기능 위험도 (인쇄 기능 포함) | **Low–Medium** → **개선 반영 후 Low** (물리 Print 제거·SaveAs 기본) |

### 개선 반영 (2026-08-04, 본 감사 §5 전 단계)

| 감사 항목 | 조치 |
|-----------|------|
| §3.1 CreateAction Print Execute | **제거** — PrintToPDFEx / RunToPDF 만 |
| §3.2 PDF 엔진 분기 | 기본 `saveas_first`, 옵션 `print_to_pdf_ex_first` |
| §3.3 모아찍기 잔존 | Print 경로·리셋 유지 + 모드 선택 |
| §3.4 지연·취소 | 시도 상한, `cancel_check`, delay 0.5s |
| §3.5 export_method | CSV/JSON·task 필드 추가 |
| §3.6 PDF 신선도 | `%PDF` 매직 + 불완전 파일 정리 |
| §3.8 캐시 신규 파일 | 폴더 mtime·파일 수 감지 |
| §3.9 문서 드리프트 | Claude.md / README / 스모크 동기화 |

### 후속 정합 (2026-08-05)

| 항목 | 조치 |
|------|------|
| SOLID 패키지 분할 | `hwp_converter/` · `hwp_print_settings/` · `windows_integration/` 등 (공개 import re-export) |
| 문서 경로 표기 | Claude.md / gemini.md / PROJECT_STRUCTURE 패키지 경로로 갱신 |
| `hwp_converter.spec` | 핵심 패키지 `hiddenimports` 명시 |
| `.gitignore` | 로컬 설정·로그·스모크 출력·일회성 분할 스크립트 패턴 |

---

## 1. Executive Summary

HwpMate는 한컴 한글 COM으로 HWP/HWPX를 PDF·문서·이미지로 일괄 변환하는 Windows 전용 PyQt6 앱이다. 단일 인스턴스, 소유 PID 강제 종료, 보안 모듈 SHA-256, SaveAs 2→3 폴백, `is_planning` 재진입 가드, 결과 원자 저장 등 **기존 코어 안전장치는 구현·테스트 상태가 양호**하다.

이번 감사의 중심은 최근 추가된 **인쇄 설정 best-effort 초기화 + PDF `PrintToPDFEx` 우선 경로**이다.

| 영역 | 평가 |
|------|------|
| 의도 | 문서에 남은 모아찍기 등이 PDF에 반영되는 문제 완화 — 방향은 타당 |
| Claude.md SaveAs 폴백 | 실패 시 2→3 인자 유지 — **구조상 유지됨** |
| 원본 디스크 비저장 | `Clear`만 사용 — **일치** |
| 신규 리스크 | **물리 프린터 `Print` Execute 폴백**, PDF 엔진 분기 품질 차이, 다단 시도 지연/취소 반응, 감사 필드 부재, 실기 스모크 부재 |
| 자동 검증 | 단위·pyright는 통과하나 **실기 한글 COM 스모크는 본 감사에서 미실행** |

**핵심 문제 요약 (신규 기능 중심):**

1. **High:** `CreateAction("Print").Execute` 폴백이 Device/가상프린터 설정 실패 시 **물리 프린터로 출력**할 수 있는 경로가 코드에 존재한다.  
2. **High/Medium:** PDF가 `SaveAs`가 아닌 `PrintToPDFEx`/가상 프린터로 먼저 나가면 **용지·품질이 SaveAs 경로와 달라질 수 있다** (비규격 용지 특히).  
3. **Medium:** 파일당 다중 COM 시도(프린터 후보 × 메서드)로 **변환 지연·취소 지연·대화상자 노출** 가능.  
4. **Medium:** `apply_default_print_settings`는 Execute 없이 속성만 건드려, SaveAs 폴백 시 **모아찍기가 그대로 남을 수 있다** (커뮤니티/한컴 포럼에서도 “속성만으로는 부족” 사례).  
5. **Low–Medium:** 결과 CSV의 `save_format`이 항상 `PDF`라 **실제 내보내기 경로 감사가 불가능**.  
6. **문서 드리프트:** 감사 당시 `Claude.md`에 `hwp_print_settings` 미기재였으나 **2026-08-04~05 동기화·패키지 경로 반영으로 해소**.

과거 감사의 High(계획 중 종료, Preflight 상한, Open=0 판정, relative_to 폴백 등)는 **상당 부분 코드에 반영된 상태**로 확인했다. 잔여는 §3 하단과 §4에 정리한다.

---

## 2. Project Understanding

### 2.1 목적과 규칙 (README / Claude.md)

| 항목 | 내용 |
|------|------|
| 목적 | HWP/HWPX → PDF·문서·이미지 일괄 변환 |
| 스택 | PyQt6, pywin32 COM, PyInstaller (`HWP변환기_v9.0.exe`, `uac_admin=True`) |
| 환경 | Windows 10/11, Python 3.10+, 한컴 한글 2018+, 관리자 권한 사실상 필수 |
| 깨지면 안 되는 로직 | SaveAs 2→3 폴백, 워커 COM apartment, 보안 모듈 실패 가시화, 네이티브 DnD, 백업 실패 비중단, artifact 성공 판정, 동일 형식 건너뜀, 단일 인스턴스, **소유 PID만** 강제 종료, `is_planning` 재진입 차단 |

README는 PDF·이미지에 대해 **「1쪽씩 best-effort, 문서 용지 유지, 환경에 따라 잔존 가능」** 이라고 안내한다. 과장 없는 수준으로 구현 의와 대체로 맞다.

### 2.2 아키텍처 (CodeGraph + 패키지)

```text
hwptopdf-hwpx_v4.py
  └─ hwpmate.app.main
       ├─ pywin32 · is_admin · SingleInstanceLock
       └─ MainWindow
            ├─ FileSelectionController  (스캔·캐시)
            ├─ ConversionController     (계획·워커·보안 세션 폴링)
            └─ ConversionWorker (QThread)
                 └─ HWPConverter.convert_file
                      ├─ Open
                      ├─ apply_default_print_settings   [신규, PDF/이미지]
                      ├─ try_export_pdf_via_print_to_pdf_ex  [신규, PDF only]
                      └─ SaveAs 2-param → 3-param 폴백
```

**CodeGraph blast radius (인쇄 기능):**

| 심볼 | 호출/의존 | 테스트 |
|------|-----------|--------|
| `apply_default_print_settings` | `convert_file` 등 | `tests/test_hwp_print_settings.py` |
| `try_export_pdf_via_print_to_pdf_ex` | `convert_file` (PDF) | 단위 + monkeypatch 통합 |
| `convert_file` | `ConversionWorker`, `tools/hwp_com_smoke.py` | `test_hwp_converter` (fake COM) |
| `ConversionWorker` | UI 컨트롤러 | `test_conversion_worker` |

### 2.3 신규 변환 흐름 (PDF)

```text
Open(forceopen)
  → sleep(DOCUMENT_LOAD_DELAY=1.0s)
  → (PDF/이미지) apply_default_print_settings  # Execute 없음, best-effort
  → before artifact snapshot
  → (PDF) try_export_pdf_via_print_to_pdf_ex
        · HAction PrintToPDFEx × 프린터 후보
        · CreateAction("Print") + Device=3 × 후보
        · XHwpPrint.RunToPDF × 후보
        · 성공 판정: 출력 파일 size>0 및 mtime/size 갱신
  → 실패 시 SaveAs(format) → SaveAs(format, "")
  → after artifact · changed 판정 · Clear
```

이미지(PNG/JPG/…)는 **리셋만** 하고 SaveAs 경로를 유지한다. DOCX 등은 인쇄 모듈을 타지 않는다 (`uses_print_settings_control`).

### 2.4 문서 vs 구현 정합

| 문서 주장 | 상태 |
|-----------|------|
| README: PDF 1쪽씩 best-effort, 용지 유지 | **대체로 일치** (강제 A4 없음) |
| Claude.md: SaveAs 2→3 폴백 유지 | **일치** (PrintToPDFEx 실패 후 폴백) |
| Claude.md: 성공 = 산출물 생성·갱신·size>0 | **일치** (경로 무관하게 artifact 검증) |
| Claude.md §3 서비스 목록 | **해소(2026-08-05)** — `hwp_print_settings/`·`hwp_converter/` 패키지 경로 반영 |
| Claude.md §2 성공 판정 서술 | **부분 불일치** — “SaveAs 폴백 후”만 명시, PrintToPDFEx 선행 미기재 |
| Claude.md Spec Kit `specs/001-...` | **불일치** — 워크스페이스에 `specs/` 없을 수 있음 (문서 자체도 가능성 명시) |
| 결과 CSV `save_format` | PrintToPDFEx 성공 시에도 `"PDF"` — **경로 구분 없음** |
| update_history / HWP_COM_SMOKE 체크리스트 | 인쇄 기능·모아찍기 샘플 항목 **반영됨** |

---

## 3. High-Risk Issues

> 실제 코드 근거가 있는 문제만 기술. 추정은 명시.

### 3.1 [신규] `CreateAction("Print").Execute` 물리 인쇄 가능성

* **위치(감사 당시):** `hwpmate/services/hwp_print_settings.py` — `try_export_pdf_via_print_to_pdf_ex` (CreateAction `"Print"` 분기).
  **현행 경로:** `hwpmate/services/hwp_print_settings/pdf_export.py` (물리 Print Execute는 이후 제거됨 — 2026-08-04 이력).
* **문제:** `PrintToPDFEx` 실패 후 `CreateAction("Print")`로 `Device=3`, `FileName`, `PrinterName`을 넣고 **`act.Execute(pset)`** 한다. 한글 버전·드라이버에 따라 Device/가상 프린터 무시 시 **기본 물리 프린터로 실제 인쇄 작업이 나갈 수 있다.** 일괄 변환 시 파일 수만큼 반복될 수 있다.
* **영향:** 용지/토너 낭비, 네트워크 프린터 대량 출력, 사용자 신뢰 훼손. **발생 빈도는 환경 의존(추정)이나 영향 심각도는 Critical에 가깝다.**
* **근거:**
  * `apply_default_print_settings` docstring은 “Execute 하지 않는다”고 명시하나, `try_export_pdf_via_print_to_pdf_ex`의 Print 분기는 **Execute 한다**.
  * 성공 판정은 “목표 경로 파일 갱신”뿐이라, 물리 인쇄가 나가도 파일이 없으면 False로 넘어가 **SaveAs까지 계속**될 수 있다(이중 부작용).
* **권장 수정 방향:**
  1. **즉시:** `CreateAction("Print").Execute` 폴백 **제거** 또는 명시적 옵트인 플래그 뒤에서만 허용  
  2. PDF 전용은 `PrintToPDFEx` + `RunToPDF` + `SaveAs` 로 한정  
  3. 실기 스모크에서 기본 프린터가 가상 PDF가 아닌 환경 검증
* **우선순위:** **Critical / High** (부작용 성격상 Critical에 가깝고, 재현은 환경 의존 → 보수적으로 **High**로 표기, 수정 우선순위는 1단계)

### 3.2 [신규] PDF 엔진 분기 — Print 경로 성공 시 레이아웃·용지 품질 차이

* **위치:** `HWPConverter.convert_file` (PDF 시 PrintToPDFEx 우선); `try_export_pdf_via_print_to_pdf_ex`
* **문제:** 기존 사용자는 `SaveAs(..., "PDF")` 품질에 맞춰져 있을 수 있다. 이제 Hancom PDF / Microsoft Print to PDF 가상 프린터가 먼저 성공하면 **폰트 임베딩·여백·비규격 용지 처리**가 달라질 수 있다. 한컴 포럼에도 가상 프린터 시 비규격 사이즈 왜곡 보고가 있다.
* **영향:** “변환은 성공”이지만 페이지 크기/잘림이 바뀌는 **조용한 품질 회귀**. 모아찍기 완화 이득과 트레이드오프.
* **근거:** `convert_file`에서 `try_export_pdf_via_print_to_pdf_ex` 성공 시 SaveAs 스킵; 프린터 후보 고정 리스트.
* **권장 수정 방향:**
  1. 기본을 **SaveAs 우선 + PrintMethod 리셋(또는 문서 인쇄설정 리셋 강화)** 로 되돌리거나  
  2. PrintToPDFEx는 **SaveAs 실패 시** 또는 **설정 옵션**으로만  
  3. 비규격 용지 샘플 스모크 추가
* **우선순위:** **High** (품질 회귀 성격)

### 3.3 [신규] 리셋 미흡 시 SaveAs 폴백에서 모아찍기 잔존

* **위치:** `apply_default_print_settings` (Execute 없음); `convert_file` SaveAs 폴백
* **문제:** 속성/`GetDefault`+SetItem만으로 세션에 PrintMethod가 고정되지 않는 환경이 커뮤니티에 보고되어 있다. PrintToPDFEx가 전부 실패하고 SaveAs로 떨어지면 **기존과 동일하게 모아찍기 PDF**가 나올 수 있다. README의 “시도” 표현과는 맞지만, 기능 목표(한쪽 인쇄) 달성률이 환경에 크게 좌우된다.
* **영향:** 기능이 “있는 것처럼 보이지만” 일부 PC에서 무효. 사용자 혼란.
* **근거:** 코드상 리셋 후 SaveAs; 한컴 포럼·pyhwpx 사례(속성만/RunToPDF 버그, PrintToPDFEx Execute로 해결).
* **권장 수정 방향:** SaveAs 직전 `PrintToPDFEx`로 **목표 파일에 PrintMethod=0 한 번 성공**하는 경로를 유지하되 §3.1 위험 폴백 제거; 또는 문서 인쇄 설정 초기화 후 SaveAs만(실기 검증).
* **우선순위:** **Medium**

### 3.4 [신규] 파일당 다중 COM 시도 — 지연·취소 둔감·UI 노출

* **위치:** `try_export_pdf_via_print_to_pdf_ex`; `DOCUMENT_LOAD_DELAY = 1.0`
* **문제:** PDF 1건당 최대 (Open sleep 1s) + (프린터 2 × 메서드 3 수준의 Execute 시도) + SaveAs. 워커는 **파일 루프 사이에서만** `cancel_requested`를 본다. 긴 Print 시도 중에는 취소가 늦게 반영된다. 일부 환경에서 인쇄 대화상자·진행 UI가 뜰 수 있다(**추정**, SetMessageBoxMode로 일부 완화).
* **영향:** 대량 PDF 배치 체감 속도 저하, 취소/종료 반응 지연, 백그라운드 작업 중 포커스 스틸.
* **근거:** 루프 구조·`ConversionWorker.run` 취소 체크 위치; `DOCUMENT_LOAD_DELAY`.
* **권장 수정 방향:** 시도 상한·조기 중단; cancel 플래그를 convert_file에 전달; PrintToPDFEx 1회만 후 즉시 SaveAs; sleep 조건부 단축.
* **우선순위:** **Medium**

### 3.5 [신규] 성공 감사 필드 부재 — 어느 경로로 뽑혔는지 불명

* **위치:** `convert_file` (`last_save_format = "PDF"` 고정); `ConversionTask.save_format`; 결과 CSV/JSON
* **문제:** PrintToPDFEx 성공과 SaveAs 성공이 결과물에서 구분되지 않는다. 품질 이슈 재현·지원이 어렵다.
* **영향:** 운영·디버깅 비용. 직접 데이터 손실은 없음.
* **근거:** `last_save_format` 대입; 모델에 export_method 필드 없음.
* **권장 수정 방향:** `last_export_method` (`print_to_pdf_ex` | `saveas_2` | `saveas_3`) 및 CSV 컬럼 추가.
* **우선순위:** **Low–Medium**

### 3.6 [신규] 산출물 신선도 판정의 약한 계약

* **위치:** `try_export_pdf_via_print_to_pdf_ex._output_looks_fresh`
* **문제:** `size>0` 및 mtime_ns 또는 size 변경만 본다. (1) 다른 프로세스가 같은 경로를 갱신하면 성공으로 오인 가능(**추정·드묾**). (2) 가상 프린터가 **다른 파일명**으로 저장하고 목표 경로를 안 건드리면 실패→SaveAs로 폴백(정상). (3) 부분 기록 후 size>0이면 성공 처리 후 깨진 PDF 가능(**추정**).
* **영향:** 드묾~중간. 후속 artifact 검증이 convert_file 바깥 before/after로 한 번 더 걸러 주나, Print 경로 내부 조기 return이 잘못된 파일을 고착시킬 수 있다.
* **근거:** 신선도 함수 구현; convert_file은 `exported=True`면 SaveAs 스킵 후 after snapshot.
* **권장 수정 방향:** PDF 매직 바이트/`%PDF` 헤더 검사; 최소 크기 임계값; 실패 시 부분 파일 삭제 후 SaveAs.
* **우선순위:** **Medium**

### 3.7 선실행 한글 attach 시 강제 종료 불가 (기존·의도)

* **위치:** `HWPConverter.initialize` (`owned_pids = after - before`); `kill_owned_processes`
* **문제:** 이미 한글이 떠 있으면 소유 PID 공집합 → 강제 종료 비활성. Claude.md 정책과 일치하는 안전장치이나 hang 시 탈출이 약하다.
* **영향:** 취소/강제 종료 UX 약화 (특히 Print 경로로 블로킹이 길어질 때 §3.4와 결합).
* **근거:** PID 차집합; Preflight 안내 존재.
* **권장 수정 방향:** 연결 전 한글 종료 권고 유지·강화; 세션 핸들 추적(신중). 전체 Hwp kill 복귀 금지.
* **우선순위:** **Medium**

### 3.8 폴더 스캔 캐시 — 신규 파일 미검출 (기존·완화됨)

* **위치:** `FileSelectionController.validate_folder_scan_cache_freshness` + TTL
* **문제:** TTL·샘플 존재 여부는 보나, 스캔 이후 **추가된 파일**은 여전히 누락될 수 있다. 샘플은 앞·뒤·중간 분산으로 개선됨.
* **영향:** 누락 변환 (완전 실패보다 조용한 누락).
* **근거:** 신선도 로직에 “신규 파일 집합 비교” 없음.
* **권장 수정 방향:** 변환 직전 경량 재스캔 옵션; 디렉터리 mtime.
* **우선순위:** **Medium** (기존과 동일 계열)

### 3.9 문서·스펙 드리프트 (기존 + 신규 가중)

* **위치:** `Claude.md` §2·§3; (없을 수 있는) `specs/`
* **문제:** 인쇄 모듈·PDF 우선 경로가 에이전트 필독 문서에 없고, Spec Kit 포인터가 비어 있을 수 있다. 이후 수정 시 SaveAs-only 가정으로 회귀할 위험.
* **영향:** 유지보수 실수, 감사 기준 불일치.
* **근거:** Claude.md 서비스 목록에 `hwp_print_settings` 없음; convert 서술 SaveAs 중심.
* **권장 수정 방향:** Claude.md에 인쇄 경로·금지 사항(물리 Print Execute)·폴백 순서 명시.
* **우선순위:** **Low–Medium**

---

## 4. Potential Functional Gaps

> “추정”은 코드만으로 단정할 수 없거나 실기 미검증 항목.

| 갭 | 설명 | 구분 |
|----|------|------|
| 이미지 모아찍기 | 이미지는 리셋만 하고 Print 전용 내보내기 없음. SaveAs 이미지가 인쇄설정을 타면 잔존 가능 | **추정** |
| A4 강제 | 사용자 초기 기대와 달리 의도적으로 미구현(원본 용지 유지). UI 옵션도 없음 | 의도적 공백 / 제품 결정 |
| 양면(Duplex) | PDF 페이지 나열과 무관; 드라이버 양면 API 미노출 | 의도적 / COM 한계 |
| 인쇄 설정 ON/OFF UI | 항상 best-effort; 끄기 옵션 없음 | 제품 갭 |
| 프린터 이름 설정 | `Hancom PDF` / `Microsoft Print to PDF` 고정 | 환경 갭 |
| 부분 PDF 정리 | Print 실패 후 0바이트·깨진 파일이 남으면 SaveAs 덮어쓰기에 의존 | 코드 갭 |
| 실기 모아찍기 스모크 | 체크리스트 항목만 있고 CI/자동화 없음 | 검증 갭 |
| export_method 감사 | §3.5 | 코드 갭 |
| 대량 계획 UI 스레드 | Preflight 상한·읽기 샘플 상한은 있음. `build_tasks` 자체는 여전히 메인 스레드 전체 순회 | 부분 완화 / 잔여 |
| Hangul 버전 매트릭스 | 2018/2020/2022/2024별 PrintToPDFEx 지원 여부 미문서화 | **추정** 필요 |

---

## 5. Recommended Fix Plan

### 1단계 — 즉시 (안전·부작용 차단)

1. **`CreateAction("Print").Execute` 폴백 제거 또는 강력 격리** (§3.1)  
2. PDF 기본 전략 재결정:  
   - **권장 A:** SaveAs 우선 + 강화된 PrintMethod 리셋  
   - **권장 B:** PrintToPDFEx **한 번만** (Hancom PDF 우선) 후 SaveAs, **Print Execute 금지**  
3. Print 실패 시 목표 경로의 **불완전 파일 삭제** 후 SaveAs  
4. `Claude.md`에 인쇄 경로·금지 사항 반영  

### 2단계 — 안정성

1. `last_export_method` + 결과 CSV/JSON 필드  
2. convert_file 내 취소 협력 / 시도 횟수 상한  
3. PDF 헤더·최소 크기 검증  
4. 모아찍기 샘플·비규격 용지 관리자 스모크 체크리스트 실행 기록  
5. DOCUMENT_LOAD_DELAY 재검토(조건 단축)  

### 3단계 — 구조·제품

1. 설정: `pdf_export_mode` (`saveas_first` | `print_to_pdf_ex_first`)  
2. 프린터 이름 사용자/자동 탐지  
3. 이미지 경로 인쇄설정 실기 검증 후 정책 확정  
4. 폴더 캐시 신규 파일 감지  
5. `PROJECT_STRUCTURE_ANALYSIS.md` / Spec 포인터 정리  

---

## 6. Test Recommendations

### 6.1 즉시 추가할 단위 테스트

| 테스트 | 목적 |
|--------|------|
| `try_export` 경로에서 `CreateAction("Print").Execute` 가 **호출되지 않음** (제거 후) 또는 플래그 off 시 미호출 | §3.1 회귀 방지 |
| PrintToPDFEx 실패 → SaveAs 2→3 순서 유지 | Claude.md 계약 |
| Print 경로가 0바이트 파일만 남긴 뒤 SaveAs가 덮어쓰기 성공 | 불완전 산출물 |
| `uses_print_settings_control` 경계 (PDF/PNG vs DOCX) | 기존 보강 |
| `last_export_method` (필드 추가 시) 기록 | 감사 필드 |

### 6.2 통합·스모크 (관리자 + 한글 실기)

| 시나리오 | 기대 |
|----------|------|
| 모아찍기(PrintMethod=4)로 저장한 HWP → PDF | 페이지 수 = 문서 쪽수(1쪽씩) |
| 일반 A4 HWP → PDF (Print vs SaveAs 각각) | 육안·페이지 박스 비교 |
| 비규격/B4 용지 → PDF | 용지 크기 유지 여부 기록 |
| 기본 프린터 = 물리 프린터, Hancom PDF 없음 | **실제 출력 0건**, SaveAs 성공 |
| 변환 중 Esc 취소 | Print 다중 시도 중에도 합리적 시간 내 중단 |
| PNG 1건 | 회귀 없음 |

### 6.3 기존 스위트

- 현재 **121 passed / pyright 0** — 유지.  
- `DOCUMENT_LOAD_DELAY` 때문에 convert_file 테스트가 느리면 전역 monkeypatch fixture 검토(기능 이슈 아님).

---

## 7. Appendix A — 이전 감사 대비 상태 (참고)

| 과거 이슈 | 현재 코드 관찰 |
|-----------|----------------|
| 계획 중 close / processEvents | **완화** — `lifecycle.close_event`가 `is_planning` 시 ignore + `close_after_plan` |
| engine_status 전 자동 클릭 | **해소** — `should_auto_accept` 가드 |
| 폴더 콜드 동기 재스캔 | **해소** — 캐시 필수 + 비동기 스캔 대기 |
| Preflight 전수 읽기 | **완화** — `PREFLIGHT_READ_CHECK_MAX_TASKS` / detail 상한 |
| Open 결과 `is False` only | **해소** — `is_com_failure_result` (False/0) |
| relative_to 미처리 | **해소** — try/except + flat 폴백 경고 |
| DLL SHA-256 | **유지** |

---

## 8. Appendix B — 검증 명령 (본 감사 시점)

```bash
python -m pytest -q
python -m pyright .
# 실기 (수동, 관리자):
# python tools/hwp_com_smoke.py --input <샘플.hwp> --format PDF --output-dir <출력>
```

---

## 9. 결론

인쇄 설정 기능은 **문제 인식과 방향(1쪽씩, 원본 용지 유지, SaveAs 폴백 유지)은 올바르다.**

**2026-08-04 개선 반영 후:** 물리 `Print` Execute 제거, 기본 `saveas_first`, `export_method` 감사, PDF 매직 검증, 취소 협력, `pdf_export_mode` 설정, 폴더 캐시 mtime 감지가 코드에 들어갔다. 잔여 리스크는 실기 한글 COM 스모크(모아찍기·비규격 용지·물리 프린터 0건)와 환경별 PrintToPDFEx 지원 여부이다.
