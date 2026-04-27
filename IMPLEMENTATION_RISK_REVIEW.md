# HwpMate 구현 리스크 개선 완료 보고서

- 작성일: 2026-04-27
- 대상 버전: v8.7
- 기준 문서: `README.md`, `claude.md`, `gemini.md`, `PROJECT_STRUCTURE_ANALYSIS.md`, `update_history.md`

## 1. 반영 완료 요약

`README.md`와 `claude.md` 기준으로 점검했던 기능 구현 리스크를 v8.7 작업으로 반영했습니다.

| 기존 리스크 | 처리 상태 | 반영 내용 |
|---|---|---|
| Python 3.9+ 표기와 실제 문법 불일치 | 완료 | 공식 지원 범위를 Python 3.10+로 정리하고 README, pyright, 가이드 문서를 갱신 |
| 하위 `backup/` 폴더 재변환 가능성 | 완료 | 폴더 재귀 스캔에서 하위 `backup/` 폴더 기본 제외 |
| `Open`/`SaveAs` 반환값과 출력 파일 미검증 | 완료 | `False` 반환 실패 처리, 출력 파일 존재 및 0바이트 초과 검증 추가 |
| COM 초기화 실패 시 결과 리포트 부재 | 완료 | 전체 실행 대상을 실패로 집계한 `ConversionSummary` 생성 |
| 동일 형식만 선택 시 결과 리포트 부재 | 완료 | 워커 실행 없이 `건너뜀` 전용 결과 다이얼로그 표시 |
| 출력 폴더 버튼 상태 불일치 | 완료 | `same_location=True` 상태에서 변환 후에도 출력 경로 UI 비활성 유지 |
| 설정 저장 원자성 부족 | 완료 | 임시 파일 저장 후 `replace()` 교체 및 손상 JSON 타임스탬프 백업 |

## 2. 추가 구현

- `AppConfig.config_version`을 `2`로 올리고 `backup_enabled`, `retry_count` 설정을 추가했습니다.
- 변환 옵션 UI에 원본 백업 on/off와 실패 시 재시도 횟수(0~3회, 기본 1회)를 추가했습니다.
- 변환 결과 CSV/JSON에 `retry_count`, `backup_file`, `backup_error` 필드를 추가했습니다.
- 사전 점검 다이얼로그에 파일별 입력 상태, 출력 경로, 충돌 조정, 백업 설정, HWP COM ProgID 감지 결과를 표시합니다.
- 출력 충돌 타임스탬프 폴백은 마이크로초와 중복 확인 루프로 보강했습니다.
- 출력 폴더 쓰기 권한 검사는 임시 파일 방식으로 변경했습니다.
- 실제 한컴오피스 COM 검증을 위한 `HWP_COM_SMOKE_TEST_CHECKLIST.md`를 추가했습니다.

## 3. 검증 결과

아래 검증을 통과했습니다.

```powershell
pyright .
pytest
pyinstaller hwp_converter.spec
```

- `pytest`: 43개 테스트 통과
- 빌드 산출물: `dist/HWP변환기_v8.7.exe`

## 4. 남은 수동 확인

자동 테스트는 실제 한컴오피스 COM 저장 동작을 대체할 수 없습니다. 배포 전에는 `HWP_COM_SMOKE_TEST_CHECKLIST.md`를 기준으로 관리자 권한 실행, HWP/HWPX 변환, 이미지 변환, 취소/강제 종료, 결과 저장을 수동 확인해야 합니다.
