# HwpMate MCP 클라이언트 설정

`python -m hwpmate.mcp`가 설치된 Windows PC에서 stdio로 연결한다.
경로는 설치 환경에 맞게 바꾼다.

## Claude Code / mcpServers 형식 클라이언트

```json
{
  "mcpServers": {
    "hwpmate": {
      "command": "C:\\Python314\\python.exe",
      "args": ["-m", "hwpmate.mcp"],
      "cwd": "D:\\twbeatles-repos\\HwpMate"
    }
  }
}
```

## Codex (TOML 형식)

Codex는 별도 TOML 설정 구조를 사용하므로 버전에 맞는 공식 설정을 확인하고,
위와 동등하게 `python -m hwpmate.mcp`를 stdio 서버 명령으로 등록한다.

## 연결 확인

1. 도구 목록에 `hwpmate_` 접두사 8개가 보이는지 확인한다.
2. `hwpmate_get_capabilities`를 호출해 `cli_path_valid`·`supported_formats`를 확인한다.
3. `cli_path_valid=false`이면 `paths.cli_executable`을 실제 설치 경로로 수정한다.
   `cli_requires_elevation=true`이면 exe가 관리자 권한을 요구하는 상태이므로, 소스 진입점
   `hwptopdf-hwpx_v4.py`를 지정하거나 호스트를 관리자 권한으로 실행한다([MCP.md](MCP.md#실행-권한-중요)).
4. 변환이 끝난 뒤 PDF 내용을 모델에 직접 반환하지 않고,
   검증 결과·파일 경로·건수·오류를 보고한다.

## 배포 메모 (PyInstaller)

- GUI용 `--windowed` EXE를 MCP 서버로 겸용하지 않는다. MCP bridge는 별도
  콘솔 바이너리로 패키징한다.
- 배포 HwpMate EXE는 `uac_admin=True`(관리자 권한 매니페스트)라 일반 권한 MCP 프로세스에서
  `subprocess`로 실행할 수 없다(`ELEVATION_REQUIRED`). exe를 변환 실행기로 쓰려면 호스트를
  관리자 권한으로 실행하거나, `asInvoker` 매니페스트의 별도 콘솔 CLI 빌드를 만든다.
- 기존 `hwp_converter.spec`의 GUI 빌드 구성은 변경하지 않는다. MCP bridge용
  spec은 `python -m hwpmate.mcp`를 진입점으로 하는 console 빌드로 별도 작성한다.
