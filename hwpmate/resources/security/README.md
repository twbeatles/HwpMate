# 한컴 한글 오토메이션 보안승인 모듈

- 파일: `FilePathCheckerModuleExample.dll`
- 출처: 한글과컴퓨터 개발자료실 오토메이션용 보안모듈  
  (미러: [hancom-io/devcenter-archive](https://github.com/hancom-io/devcenter-archive) `hwp-automation/보안모듈(Automation).zip`)
- 용도: COM 자동화 시 파일 Open/Save 경로 승인 팝업 억제
- 런타임: 앱이 `%LOCALAPPDATA%\HwpMate\security\` 로 복사 후  
  `HKCU\SOFTWARE\HNC\HwpAutomation\Modules` 에 경로를 등록하고  
  `RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")` 를 호출합니다.

한컴 공개 자료이며, 임의 개작·재배포 시 한컴 라이선스를 확인하세요.
