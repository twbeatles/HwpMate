# Project Audit

## 1. Executive Summary

2026-06-11 감사에서 확인한 기능 구현 리스크는 코드로 보강 완료했다. 주요 변경은 변환 중 중복 실행 차단, 이미지/HTML 보조 산출물 충돌 회피, HWP `Open(False)` 실패 경로 정리, 설정/결과 저장 실패 처리, 실패 상태 표시 보존, 단일 인스턴스 잠금, Windows 경로 검증 강화다.

현재 자동검증 기준 위험도는 **Low-Medium**이다. 일반 단위/정적 검증으로 재현 가능한 리스크는 회귀 테스트를 추가해 닫았고, 남은 위험은 실제 한컴오피스 COM, 관리자 권한 GUI, PyInstaller 산출물 실행처럼 로컬 환경 의존적인 수동/릴리스 검증 영역이다.

검증 결과:

- `python -m pytest -q` : 67 passed
- `python -m pyright .` : 0 errors, 0 warnings, 0 informations

## 2. Project Understanding

HwpMate는 `hwptopdf-hwpx_v4.py`에서 `hwpmate.bootstrap.main`을 호출하는 Windows 전용 PyQt6/pywin32 GUI 도구다. README.md와 `claude.md` 기준 유지보수 대상은 `hwpmate/` 패키지이며, `legacy/hwptopdf-hwpx v3.py`는 참고용이다.

CodeGraph 분석 기준 주요 흐름은 다음과 같다.

1. `hwpmate.app.main()`이 pywin32, 관리자 권한, 단일 인스턴스 잠금을 확인한 뒤 `MainWindow`를 생성한다.
2. `MainWindow`는 공개 import 경로를 유지하는 조립 루트이며, 실제 책임은 `AppearanceController`, `FileSelectionController`, `ConversionController`, `NativeDropController`, `LifecycleController`로 분리된다.
3. 파일/폴더 입력은 `FileScanWorker`와 `FileSelectionStore`를 거쳐 변환 계획으로 이어진다.
4. `TaskPlanner`가 실행/건너뜀/출력 경로를 계산하고, `artifact_policy`를 통해 기본 출력 파일과 이미지/HTML 보조 산출물 충돌을 함께 회피한다.
5. `ConversionWorker`가 워커 스레드에서 COM 초기화, 백업, 변환/재시도, 취소, summary emit, cleanup을 수행한다.
6. `HWPConverter.convert_file()`은 `Open`, `SaveAs` 2-인자/3-인자 폴백, 산출물 snapshot 비교, `Clear` 정리를 수행한다.
7. `ResultDialog`는 실패 TXT와 전체 결과 CSV/JSON을 원자 저장 방식으로 기록한다.

## 3. High-Risk Issues

### 3.1 변환 중 명령 진입점 가드 부족

위치: `ConversionController.start_conversion`, `FileSelectionController`, `NativeDropController`, `LifecycleController`, `AppearanceController`

문제: 버튼 비활성화만으로는 `Ctrl+Enter`, 메뉴 액션, 네이티브 드롭이 변환 중 상태를 우회할 수 있었다.

조치: `start_conversion()`에 실행 중 worker/busy guard를 추가했고, 파일 추가/삭제, 폴더/출력 변경, 네이티브 드롭 진입점도 변환 중에는 상태를 바꾸지 않도록 차단했다. 메뉴 `QAction`과 시작 `QShortcut`은 컨트롤러 속성으로 보관해 변환 중 비활성화한다.

검증: 변환 중 중복 시작, 파일 목록 변경, 네이티브 드롭, 메뉴/단축키 비활성 회귀 테스트 추가.

우선순위: High -> **완료**

### 3.2 이미지/HTML 보조 산출물 충돌 회피 누락

위치: `hwpmate/services/artifact_policy.py`, `TaskPlanner.resolve_output_conflicts`, `HWPConverter._iter_candidate_artifact_files`

문제: 성공 판정은 same-stem 보조 산출물을 보지만, 충돌 회피는 기본 출력 파일만 확인했다.

조치: `artifact_policy`를 추가해 보조 산출물 후보, existing conflict, snapshot 대상 규칙을 공유한다. PNG/JPG/BMP/GIF/HTML은 기본 출력 파일이 없어도 `doc_001.png`, `doc.files` 같은 same-stem 보조 산출물이 있으면 덮어쓰기 해제 상태에서 새 이름으로 리네임된다. 접두 false positive를 줄이기 위해 delimiter 경계 매칭을 적용하고, 보조 디렉터리 재귀 snapshot에는 상한을 둔다.

검증: 기존 보조 산출물 충돌 리네임, unrelated prefix 무시, nested scan limit 테스트 추가.

우선순위: High -> **완료**

### 3.3 HWP Open False 경로 정리 누락

위치: `HWPConverter.convert_file`

문제: `Open()`이 `False`를 반환하는 실패 경로에서 `Clear()`가 호출되지 않았다.

조치: Open False 반환 시 best-effort `hwp.Clear(option=1)`를 수행하도록 보강했다.

검증: Open False 실패 테스트가 `Clear(option=1)` 호출까지 확인한다.

우선순위: Medium -> **완료**

### 3.4 설정 저장 실패의 UI 미전파

위치: `ConfigRepository.save`, `save_config`, `LifecycleController.save_settings`, `AppearanceController.toggle_theme`

문제: 설정 저장 실패가 로그에만 남고 호출자/UI로 전파되지 않았다.

조치: `ConfigRepository.save()`와 `save_config()`가 bool을 반환한다. 종료/테마 변경 저장 실패 시 status/toast 또는 종료 경고로 사용자에게 알린다. 기존 호출자 호환성을 위해 `False`만 명시 실패로 처리한다.

검증: 설정 저장 성공/실패 반환 테스트와 UI 경로 회귀 테스트 추가.

우선순위: Medium -> **완료**

### 3.5 실패 직후 HWP 상태 표시 덮어쓰기

위치: `ConversionController.on_task_completed/on_worker_finished`

문제: 실패 summary가 설정한 오류 상태가 worker finished 처리에서 항상 정상 대기로 덮였다.

조치: 마지막 summary에 실패나 취소가 있으면 HWP 상태 라벨을 각각 실패/취소 상태로 유지한다.

검증: 실패 summary 이후 worker finished에서도 상태 라벨이 실패를 유지하는 테스트 추가.

우선순위: Low -> **완료**

## 4. Potential Functional Gaps

- **단일 인스턴스 잠금**: `hwpmate/app_instance.py`의 `QLockFile` 기반 `SingleInstanceLock`으로 구현 완료. 두 번째 실행은 안내 후 종료한다.
- **Windows 경로 검증**: 예약 장치명, trailing dot/space, 제어 문자, 잘못된 colon 사용을 거부하도록 `is_valid_path_name()`을 강화했다.
- **결과 저장 원자성**: 실패 TXT, 결과 CSV, 결과 JSON 저장을 임시 파일 후 replace 방식으로 변경했다.
- **실제 COM/빌드 검증**: 이번 작업 범위는 자동검증으로 한정했다. 실제 HWP COM 변환과 PyInstaller 실행 파일 검증은 `HWP_COM_SMOKE_TEST_CHECKLIST.md`의 릴리스 체크 항목으로 남는다.

## 5. Recommended Fix Plan

### 1단계: 즉시 수정

- 변환 중 중복 시작/입력 변경/네이티브 드롭 차단 완료.
- 이미지/HTML 보조 산출물 충돌 회피 완료.

### 2단계: 안정성 개선

- Open False 정리, 설정 저장 실패 전파, 실패 상태 표시 보존, 결과 리포트 원자 저장 완료.

### 3단계: 구조 개선

- `artifact_policy`와 `SingleInstanceLock` 내부 모듈 추가 완료.
- README, `claude.md`, `HWP_COM_SMOKE_TEST_CHECKLIST.md`, `update_history.md`, `hwp_converter.spec` 주석을 현재 구현 기준으로 동기화 완료.

## 6. Test Recommendations

자동화 완료:

- 변환 중 중복 시작, 메뉴/단축키 비활성, 파일 목록 변경 차단, 네이티브 드롭 차단
- 이미지/HTML 보조 산출물 충돌 리네임, 접두 false positive 방지, nested scan limit
- `Open(False)` 경로 `Clear(option=1)` 정리
- 설정 저장 성공/실패 반환
- 결과 TXT/CSV/JSON 원자 저장 후 임시 파일 정리
- 실패 summary 후 HWP 상태 표시 유지
- `QLockFile` 단일 인스턴스 잠금
- Windows 예약명/trailing dot/space/제어 문자/colon 경로 검증

수동/릴리스 검증으로 남길 항목:

- 관리자 권한 GUI 실행
- 실제 HWP/HWPX -> PDF/DOCX/이미지/HTML 변환
- 네이티브 관리자 드래그 앤 드롭
- PyInstaller `hwp_converter.spec` 빌드와 `HWP변환기_v8.7.exe` 실행
