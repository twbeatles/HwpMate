# Project Audit

**최종 갱신:** 2026-08-05  
**대상:** HwpMate (`hwptopdf-hwpx_v4.py` → `hwpmate/`)  
**문서 정합:** README · Claude.md · gemini.md · PROJECT_STRUCTURE · 스모크 체크리스트 · `hwp_converter.spec` · `.gitignore`

| 검증 | 결과 |
|------|------|
| `python -m pytest -q` | **143 passed** |
| `python -m pyright .` | **0 errors / 0 warnings** |
| 전체 기능 위험도 | **Low** |

## 반영된 개선 (2026-08-05)

| 영역 | 내용 |
|------|------|
| 보안·세션 | UI 폴링 + 워커 suppress 모두 **소유 PID만** HWND 조작; 자동 클릭은 모듈 실패 시에만 |
| PDF | SaveAs·Print 모두 `%PDF` 매직; SaveAs 실패/무산출 시 Print 폴백 |
| 재진입 | `is_planning` + `_confirm_preflight_and_start_worker` (시작·실패 재변환) |
| 취소 | COM 대기 안내, 슬라이스 wait, 강제종료 대기 연장 |
| 폴더 캐시 | 변환 직전 90초 연령, 만료 시 재스캔; mtime 단독 하드 실패 없음 |
| 경로 | `com_path_candidates` 확장 경로 재시도; Preflight 240 경고 / 260 차단 |
| 백업 | `backup_max_files_per_stem` UI·설정·prune |
| PDF UI | 내보내기 모드 콤보 |
| 암호 힌트 | Open/예외 메시지 best-effort |

## 잔여 리스크 (의도적·환경)

- COM 블로킹 중 즉시 취소 불가 (안내·강제종료로 완화)
- 실기 한글 COM CI 없음 → `HWP_COM_SMOKE_TEST_CHECKLIST.md` / `tools/hwp_com_smoke.py`
- 한컴 버전별 인쇄 설정 best-effort 한계

## 품질 게이트

```bash
pyright .
pytest
pyinstaller --noconfirm --clean hwp_converter.spec
# 가능 시 관리자 권한
python tools/hwp_com_smoke.py --input <샘플.hwp> --format PDF --output-dir <출력>
```
