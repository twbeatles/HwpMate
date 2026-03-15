# HwpMate 프로젝트 지침서 (Gemini)

이 문서는 Gemini 계열 에이전트가 HwpMate를 유지보수할 때 필요한 핵심 맥락을 요약합니다. 현재 기준 구현은 `hwptopdf-hwpx_v4.py`이며, 모든 문서와 빌드는 이 파일을 우선 기준으로 봅니다.

## 1. 핵심 사실

- 앱 목적: HWP/HWPX 문서를 여러 출력 형식으로 일괄 변환
- 플랫폼: Windows 전용
- GUI: `PyQt6`
- 자동화: `pywin32` 기반 HWP COM
- 빌드: `PyInstaller` + `hwp_converter.spec`
- 레거시 참고 파일: `hwptopdf-hwpx v3.py`

## 2. 꼭 유지해야 하는 동작

### 변환 엔진 호환성
- `Open` 이후 짧은 대기와 `SaveAs` 폴백은 실제 한글 버전 차이를 흡수하기 위한 장치입니다.
- 출력 형식 정보는 `FORMAT_TYPES`를 기준으로 관리합니다.

### 관리자 권한 드롭 처리
- 앱은 관리자 권한을 전제로 동작합니다.
- 드래그 앤 드롭은 `NativeDropFilter`의 Windows 메시지 처리까지 포함해야 완전합니다.

### 워커 분리
- 파일 스캔과 변환이 UI 스레드를 막지 않도록 `FileScanWorker`, `ConversionWorker`를 유지합니다.
- 스레드 내부 COM 초기화는 삭제하지 않습니다.

### 백업과 덮어쓰기 방지
- `_create_backup`은 원본 보호를 위한 기본 안전장치입니다.
- 덮어쓰기 미허용 시 자동 번호 부여 로직을 유지합니다.

## 3. 현재 리포지토리 품질 기준

- `pyright .` 가 0 오류여야 합니다.
- `pyrightconfig.json`을 기준으로 타입 검사를 맞춥니다.
- `.editorconfig` 기준으로 `utf-8`, `LF`, 최종 개행을 유지합니다.
- 문서와 실제 코드가 어긋나면 코드 변경과 함께 문서도 같이 고칩니다.

## 4. 파일별 역할

- `hwptopdf-hwpx_v4.py`: 운영 코드 전체
- `hwp_converter.spec`: PyInstaller 빌드 정의
- `README.md`: 사용자 안내와 실행/빌드 방법
- `update_history.md`: 기능/유지보수 이력
- `claude.md`, `gemini.md`: 협업 에이전트용 개발 가이드

## 5. 변경 시 체크 포인트

1. 지원 형식을 바꾸면 `FORMAT_TYPES`, README, 히스토리를 같이 갱신합니다.
2. 빌드 이름이나 배포 방식을 바꾸면 `.spec`과 README를 같이 갱신합니다.
3. 타입 관련 수정 후에는 반드시 `pyright .`를 다시 실행합니다.
4. 변환 중 생성되는 `backup/` 폴더 같은 산출물은 Git 추적 대상이 아니어야 합니다.
5. 관리자 권한 요구사항을 약화시키는 변경은 신중하게 검토합니다.

## 6. 권장 검증 명령

```bash
pyright .
pyinstaller hwp_converter.spec
```

가능하면 추가로 수동 확인할 것:

- 앱 실행
- 폴더 변환 1회
- 파일 드롭 1회
- 결과 폴더 열기와 실패 목록 저장 기능
