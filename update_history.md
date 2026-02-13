HWP to PDF/HWPX 변환기 개발 히스토리

이 문서는 Python과 pywin32를 활용한 한글(HWP) 자동화 변환기 프로젝트의 버전별 변경 사항과 기술적 이슈 해결 과정을 기록합니다.

✅ 최종 버전: v8.3

안정화 및 호환성 최종 수정

주요 수정: 한글 버전에 따라 SaveAs 메소드의 인자 개수가 달라 발생하는 TypeError 해결.

세부 내용:

hwp.SaveAs(path, format, "") 형태로 3번째 인자(빈 문자열)를 명시적으로 전달하여 호환성 확보.

파일 열기(Open)와 저장(SaveAs) 사이 안정적인 처리를 위해 time.sleep(1.0) 대기 시간 유지.

👑 v8.6 (Modern UI & Stability)

통합 및 안정성 강화:

코드베이스 동기화: 개발 버전(dist copy)을 메인으로 통합 및 정리.

로깅 시스템: RotatingFileHandler 도입 (로그 자동 순환, 시스템 정보 기록).

프로세스 안전장치: 변환 중 '응답 없음' 시 좀비 프로세스(Hwp.exe) 강제 종료(Force Kill) 기능 추가.

최적화:

대량 파일 추가 시 UI 프리징 현상 해결 (SetUpdatesEnabled 최적화).

결과 창에서 '폴더 열기' 시 해당 파일 자동 선택(Highlight) 기능 지원.

🛠 v8.0 ~ v8.2: 구조 개선 및 레지스트리 오류 대응

리팩토링 및 COM 객체 연결 안정화

v8.2 (Security Module Fix)

버그 수정: v8.1에서 누락되었던 보안 모듈(FilePathCheckDLL) 등록 로직 복구.

안정성: 문서 로딩 시점 확보를 위한 대기 로직(time.sleep) 추가.

v8.1 (Registry Fix)

이슈: "잘못된 클래스 문자열입니다" 오류(COM 레지스트리 꼬임) 재발.

해결: 3단계 연결 시도 로직 구현.

win32com.client.dynamic.Dispatch (캐시 무시)

win32com.client.Dispatch (일반 연결)

gen_py (캐시 폴더) 강제 삭제 후 재시도

v8.0 (Refactoring)

구조 변경: 기존 단일 클래스 구조를 MVC 패턴과 유사하게 분리.

HwpLogic: 한글 제어 및 변환 로직 담당.

HWPConverterUI: tkinter 화면 구성 및 이벤트 처리 담당.

코드 정리: 초기화 로직을 단순화하여 win32.Dispatch 원복 시도 (이후 v8.1에서 보강).

🚀 v7.0 ~ v7.2: 편의 기능 및 자동 복구

드래그 앤 드롭 및 캐시 자동 삭제

v7.2 (Dynamic Dispatch)

기술 변경: gencache 방식 실패 시 dynamic.Dispatch를 사용하는 폴백(Fallback) 로직 추가.

오류 리포트: traceback 모듈을 사용하여 오류 발생 시 상세 스택 트레이스 출력 기능 추가.

v7.1 (Auto-Fix)

기능 추가: AttributeError: CLSIDToClassMap 등 캐시 손상 오류 감지 시, 자동으로 gen_py 폴더를 찾아 삭제하는 기능 추가.

v7.0 (DND Support)

기능 추가: tkinterdnd2 라이브러리를 활용한 파일 드래그 앤 드롭 지원.

예외 처리: 라이브러리 미설치 시에도 기본 버튼 기능은 동작하도록 예외 처리 적용.

✨ v4.0 ~ v6.0: 기능 확장

포맷 지원 확장 및 변환 모드 다변화

v6.0 (Stop Function)

기능 추가: 변환 작업 중단(Stop) 버튼 및 로직 추가.

v5.1 (Popup Fix)

버그 수정: 변환 종료 시 "빈 문서 1을 저장하시겠습니까?" 팝업이 뜨는 문제 해결.

해결: Quit 호출 전 hwp.Clear(3)(저장 안 하고 모두 닫기) 수행.

v5.0 (File Selection Mode)

기능 추가: 기존 '폴더 일괄 변환' 외에 '개별 파일 선택 변환' 모드 추가.

UI 변경: 변환 모드 선택 라디오 버튼 및 파일 리스트박스 UI 구현.

v4.0 (HWPX Support)

기능 추가: 출력 포맷을 PDF와 HWPX 중 선택하는 옵션 추가.

로직 변경: .hwp 파일뿐만 아니라 .hwpx 파일도 입력 대상으로 인식하도록 수정.

📦 v1.0 ~ v3.6: 초기 개발 및 기반 마련

기본 기능 구현 및 필수 옵션 추가

v3.0 ~ v3.6

기능 추가:

하위 폴더 포함(Recursive) 검색 기능.

결과 상세 리포트 창(성공/실패 목록) 구현.

기존 파일 덮어쓰기 방지 옵션.

버그 수정: 긴 파일 경로 처리 시도 및 SyntaxError 유발하는 문자열 이스케이프 문제 수정.

v2.0 ~ v2.11

초기 안정화:

관리자 권한 실행 감지 및 경고 기능.

FileSaveAsPdf 액션 오류 해결을 위해 Print 방식 시도 후, 최종적으로 SaveAs 메소드로 정착.

프로그레스바(진행률) 구현.

입/출력 폴더 경로 지정 기능 구현.

v1.0

초기 프로토타입: 지정된 폴더 내 HWP 파일을 PDF로 변환하는 기본 로직 구현.

📝 기술적 요약 (Key Learnings)

COM 객체 연결: 한글(HWP) 자동화 시 win32.Dispatch와 gencache.EnsureDispatch의 차이, 그리고 레지스트리 오류(gen_py) 발생 시 대응 방법이 핵심이었습니다.

보안 모듈: 최신 한글 버전에서는 외부 스크립트 실행 시 RegisterModule("FilePathCheckDLL", ...)이 필수적입니다.

함수 인자: SaveAs 메소드는 버전에 따라 인자 개수(2개 또는 3개)에 민감하게 반응하므로 빈 문자열("")이라도 명시하는 것이 호환성에 좋습니다.

프로세스 관리: 파이썬 스레드(threading) 내에서 COM 객체 사용 시 pythoncom.CoInitialize() 호출이 필수입니다.