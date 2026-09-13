# HwpMate — 한글(HWP/HWPX) to PDF·Word·이미지 초고속 일괄 변환기

<p align="center">
  <a href="https://github.com/twbeatles/HwpMate/releases/latest">
    <img src="https://img.shields.io/badge/Release-v9.1.0-blue.svg?style=for-the-badge&logo=github" alt="Latest Release" />
  </a>
  <img src="https://img.shields.io/badge/Platform-Windows_10%2F11_(64--bit)-lightgrey.svg?style=for-the-badge&logo=windows" alt="Platform" />
  <img src="https://img.shields.io/badge/Python-3.10+-3776AB.svg?style=for-the-badge&logo=python&logoColor=white" alt="Python 3.10+" />
  <img src="https://img.shields.io/badge/Tests-168%20Passed-success.svg?style=for-the-badge" alt="Tests" />
  <img src="https://img.shields.io/badge/Type%20Check-Pyright%200%20Errors-blueviolet.svg?style=for-the-badge" alt="Type Check" />
</p>

<p align="center">
  <strong>한컴오피스 한글 문서(HWP, HWPX)를 PDF, Word(DOCX), 이미지(PNG, JPG), TXT 등으로 단 몇 번의 클릭 또는 명령어 하나로 대량 일괄 변환하는 Windows 데스크톱 & CLI 도구입니다.</strong><br>
  <em>High-Performance Batch Converter for Hancom Hangul (.hwp, .hwpx to PDF, DOCX, Images & Text)</em>
</p>

---

> ### 🚀 [초고속 다운로드 (무설치 포터블)]
> Python 설치나 복잡한 환경 설정 없이, 아래 링크에서 단일 실행 파일(`.exe`)을 다운로드하여 즉시 사용하실 수 있습니다.
> 
> 👉 **[최신 버전 HwpMate-v9.1.0.exe 다운로드 (GitHub Releases)](https://github.com/twbeatles/HwpMate/releases/latest)**  
> *(다운로드 후 마우스 오른쪽 버튼을 클릭하여 **'관리자 권한으로 실행'**해 주세요.)*

---

## 📌 목차
- [주요 특징 (Why HwpMate?)](#features)
- [변환 지원 포맷](#supported-formats)
- [시스템 요구사항](#requirements)
- [1분 퀵스타트 (GUI 사용법)](#quick-start)
- [커맨드라인(CLI) 자동화 변환](#cli)
- [주요 단축키](#shortcuts)
- [자주 묻는 질문 및 문제 해결 (FAQ)](#faq)
- [개발자 및 기여 가이드](#developer-guide)
- [관련 프로젝트](#related-projects)

---

## <a id="features"></a>✨ 주요 특징 (Why HwpMate?)

수백~수천 개의 한글 문서를 수작업으로 일일이 열어서 PDF나 Word로 "다른 이름으로 저장"하는 번거로운 작업은 이제 그만!  
**HwpMate**는 대량 변환 시 발생하는 **한컴 보안 팝업, 메모리 누수, 프로세스 먹통, 파일 유실** 문제를 엔지니어링 수준에서 완벽하게 해결했습니다.

* ⚡ **초고속 대량 일괄 변환 (Batch Conversion)**:
  폴더 전체(하위 폴더 포함 가능) 또는 원하는 파일들을 드래그 앤 드롭하여 수백 개의 문서를 한 번에 고속 변환합니다.
* 🛡️ **한컴 보안 승인 팝업 자동 처리**:
  한컴오피스의 귀찮은 “접근 허용” 확인 창을 차단하는 보안 모듈(`FilePathChecker`)을 자동 내장 및 등록하며, 미등록 환경에서도 팝업을 자동 감지하여 창 전면화 및 자동 클릭을 지원합니다.
* 🔄 **메모리 누수 제로 (200건 단위 프로세스 자동 재순환)**:
  대량 문서 변환 시 한컴 프로세스가 느려지거나 멈추는 현상을 방지하기 위해, 200건 변환마다 백그라운드 한글 COM 인스턴스를 무결하게 재순환(Recycle)합니다.
* 💾 **철저한 원본 보호 및 안전 백업**:
  변환 작업 전 원본 문서를 `backup/` 폴더에 자동으로 안전하게 복사하며, 중복 파일 덮어쓰기 방지 및 백업 수량 순환 보관(1~100개)을 지원합니다.
* 🎯 **듀얼 PDF 엔진 탑재**:
  고품질 벡터 그래픽을 유지하는 **SaveAs(용지 품질 우선)** 모드와 인쇄 설정 왜곡을 완화하는 **PrintToPDFEx(모아찍기 해제 우선)** 모드를 자유롭게 선택할 수 있으며, 산출물 유효성 검증 실패 시 자동 교차 폴백됩니다.
* 🔒 **Ed25519 전자 서명 기반 자동 업데이트**:
  새 버전 출시 시 앱 내에서 원클릭으로 안전하게 업데이트되며, 무결성 검증 실패 시 이전 버전으로 즉시 **자동 롤백** 복구됩니다.
* 💻 **직관적인 모던 GUI & 헤드리스 CLI 지원**:
  초보자를 위한 미려하고 직관적인 GUI 화면과, 업무 자동화·배치 스크립트를 위한 터미널 CLI 명령어를 모두 제공합니다.

---

## <a id="supported-formats"></a>📂 변환 지원 포맷

`.hwp` 및 최신 표준 포맷인 `.hwpx` 문서를 다양한 형태의 문서와 이미지로 변환할 수 있습니다.

| 구분 | 지원 포맷 | 확장자 | 특징 및 주요 용도 |
| :--- | :--- | :--- | :--- |
| **문서** | **PDF** | `.pdf` | 전자결재, 공문서 배포, 인쇄용 표준 문서 (듀얼 엔진 지원) |
| | **MS Word** | `.docx` | Microsoft Office Word 호환 문서로 변환 |
| | **한글 표준** | `.hwpx` | 레거시 HWP 문서를 최신 개방형 HWPX 포맷으로 일괄 업그레이드 |
| | **한글 문서** | `.hwp` | HWPX를 기존 HWP 5.0 포맷으로 다운그레이드 변환 |
| | **웹 문서** | `.html` | 웹 브라우저 열람 및 웹 퍼블리싱용 HTML 변환 |
| | **오픈 오피스** | `.odt` | ODF 표준 오피스 문서 변환 |
| | **텍스트/서식** | `.txt`, `.rtf` | 텍스트 데이터 추출 및 서식 있는 텍스트 |
| **이미지** | **PNG / JPG** | `.png`, `.jpg` | 고화질 페이지별 이미지 추출 (웹 게시, 썸네일용) |
| | **BMP / GIF** | `.bmp`, `.gif` | 비압축 비트맵 및 웹 이미지 변환 |

> [!NOTE]
> * 변환 대상과 동일한 형식의 파일(예: PDF로 변환 시 대상 폴더에 이미 존재하는 `.pdf` 파일)은 **자동으로 건너뜁니다.**
> * 암호가 걸린 문서나 손상된 파일은 사전 점검 및 변환 단계에서 안전하게 감지되어 건너뛰고, 변환 완료 후 리포트로 안내됩니다.

---

## <a id="requirements"></a>💻 시스템 요구사항

* **운영체제**: Windows 10 또는 Windows 11 (64-bit 권장)
* **필수 소프트웨어**: **한컴오피스 한글 2018 이상** (2018, 2020, 2022, 2024 등 정식 설치 버전)
* **실행 권한**: **관리자 권한** (한글 COM 컴포넌트 제어 및 보안 모듈 레지스트리 등록에 필수)
* *배포용 실행 파일(`.exe`) 사용 시 Python 설치는 전혀 필요 없습니다.*

---

## <a id="quick-start"></a>⚡ 1분 퀵스타트 (GUI 사용법)

```
[1. 파일/폴더 추가]  ➔  [2. 변환 포맷 선택]  ➔  [3. 사전 점검 & 변환 시작]  ➔  [4. 결과 리포트]
   (드래그 앤 드롭)         (PDF / DOCX 등)             (Ctrl + Enter)          (성공/실패 즉시 확인)
```

### 1단계: 프로그램 실행
다운로드한 `HwpMate-v9.1.0.exe` 파일을 마우스 우클릭한 뒤 **[관리자 권한으로 실행]**을 클릭합니다.

### 2단계: 문서 또는 폴더 추가
* **폴더 통째로 변환할 때**: 상단 모드를 **[폴더 일괄 변환]**으로 두고, 변환할 폴더를 창으로 끌어다 놓거나 `[폴더 선택]`을 누릅니다. (하위 폴더를 포함하려면 체크박스 선택)
* **선택한 파일만 변환할 때**: 모드를 **[파일 개별 선택]**으로 두고, 파일들을 마우스로 드래그 앤 드롭하여 추가합니다.

### 3단계: 변환 포맷 선택
상단의 형식 카드에서 변환하고자 하는 목표 형식(**PDF**, **DOCX**, **PNG** 등)을 클릭합니다.

### 4단계: 변환 시작
* 우측 하단의 **[변환 시작]** 버튼을 누르거나 단축키 `Ctrl + Enter`를 누릅니다.
* **사전 점검(Preflight Check)** 팝업에서 총 변환 대상 개수, 예상 건너뜀, 저장 경로를 확인하고 [확인]을 누르면 변환이 진행됩니다.
* 변환이 완료되면 결과 창에서 성공/실패 통계를 확인하고, 실패한 파일만 원클릭으로 재시도하거나 결과 폴더를 즉시 열 수 있습니다.

---

## <a id="cli"></a>⚙️ 커맨드라인(CLI) 자동화 변환

HwpMate는 GUI 창 없이 윈도우 작업 스케줄러, 배치 파일(`.bat`), 파이썬 스크립트 등과 연동할 수 있는 강력한 **헤드리스 커맨드라인 모드**를 제공합니다.

### CLI 사용 예시

```powershell
# 1. 단일 파일 변환 (HWP -> PDF)
HwpMate-v9.1.0.exe --input "C:\문서\보고서.hwp" --format PDF

# 2. 폴더 전체 일괄 변환 (하위 폴더 포함, MS Word DOCX로 변환)
HwpMate-v9.1.0.exe --input "C:\업무자료" --format DOCX --recursive --output "C:\변환완료"

# 3. 초고속 변환 (백업 생성 안 함, 기존 파일 덮어쓰기)
HwpMate-v9.1.0.exe --input "C:\문서폴더" --format PDF --overwrite --no-backup

# 4. 모아찍기 해제 인쇄 모드로 PDF 변환 (실패 시 2회 재시도)
HwpMate-v9.1.0.exe --input "C:\공문서" --format PDF --pdf-export-mode print_to_pdf_ex_first --retry 2

# 5. 시스템 및 COM 의존성 무결성 진단 (스모크 테스트)
HwpMate-v9.1.0.exe --smoke
```

### CLI 옵션 상세 안내

| 옵션 | 단축키 | 기본값 | 설명 |
| :--- | :---: | :---: | :--- |
| `--input` | `-i` | (필수) | 변환할 HWP/HWPX 단일 파일 또는 폴더 경로 |
| `--format` | `-f` | `PDF` | 변환 목표 형식 (`PDF`, `DOCX`, `HWPX`, `HWP`, `PNG`, `JPG`, `TXT` 등) |
| `--output` | `-o` | 원본 위치 | 변환 결과물이 저장될 폴더 경로 (생략 시 원본과 같은 위치에 생성) |
| `--recursive` | `-r` | `False` | 폴더 입력 시 하위 폴더의 모든 한글 문서를 재귀적으로 탐색하여 변환 |
| `--overwrite` | | `False` | 결과 경로에 동일 이름 파일이 존재할 경우 덮어쓰기 |
| `--no-backup` | | `False` | 변환 전 원본 파일 백업(`backup/` 폴더) 생성을 건너뜀 |
| `--retry` | | `1` | 변환 실패 시 자동 재시도 횟수 (`0` ~ `3`) |
| `--pdf-export-mode` | | `saveas_first` | PDF 엔진 우선순위 (`saveas_first`: 용지 품질 / `print_to_pdf_ex_first`: 모아찍기 완화) |
| `--smoke` | | `False` | GUI를 띄우지 않고 의존성 및 모듈 무결성 점검 실행 후 종료 |

---

## <a id="shortcuts"></a>⌨️ 주요 단축키

| 단축키 | 동작 |
| :--- | :--- |
| `Ctrl + Enter` | **변환 시작** (사전 점검 창 호출) |
| `Esc` | **변환 작업 취소** |
| `Ctrl + O` | 파일 추가 대화상자 열기 |
| `Ctrl + Shift + O` | 폴더 선택 대화상자 열기 |
| `Delete` | 파일 목록에서 선택한 항목 제거 |
| `Ctrl + Delete` | 파일 목록 전체 비우기 |
| `F1` | 프로그램 정보 및 단축키 안내 |

---

## <a id="faq"></a>❓ 자주 묻는 질문 및 문제 해결 (FAQ)

### Q1. 한글 프로그램의 "접근 허용" 보안 확인 창이 계속 뜹니다.
> **A:** HwpMate는 한컴 공식 보안 모듈(`FilePathCheckerModuleExample.dll`)을 내장하고 있어, 변환 시작 시 자동으로 레지스트리에 등록하여 보안 창을 원천 차단합니다.  
> 만약 보안 창이 계속 뜬다면 프로그램을 반드시 **[관리자 권한으로 실행]**했는지 확인해 주세요. 백신 프로그램 등에 의해 임시 등록이 차단된 경우, HwpMate의 [허용 창 자동 클릭] 기능이 보조 동작하므로 화면 뒤 작업 표시줄을 확인해 주시기 바랍니다.

### Q2. 반드시 '관리자 권한'으로 실행해야 하나요?
> **A:** 네, 권장되며 사실상 필수입니다. Windows의 COM 자동화 인터페이스를 통해 백그라운드 한글 프로세스를 안정적으로 제어하고 보안 모듈 레지스트리를 연동하려면 관리자 권한이 요구됩니다.

### Q3. 한컴오피스 한글이 설치되어 있지 않은 PC나 리눅스(Linux)에서도 동작하나요?
> **A:** HwpMate는 한컴오피스 한글의 공식 COM API 엔진을 직접 호출하여 원본 문서의 표, 수식, 레이아웃, 글꼴을 100% 완벽하게 보존하며 변환합니다. 따라서 **한컴오피스 한글 2018 이상이 설치된 Windows 환경이 필수**입니다.

### Q4. PDF로 변환했을 때 2쪽 모아찍기나 용지 크기가 어긋납니다.
> **A:** 이전에 한글 프로그램에서 사용했던 인쇄 모아찍기 캐시가 남아있을 경우 발생할 수 있습니다. 변환 옵션에서 **[PDF 내보내기 모드]**를 `PrintToPDFEx 우선 (모아찍기 완화)`으로 변경하여 변환해 보세요. HwpMate는 변환 시 1쪽씩 기본 인쇄 설정으로 자동 초기화를 시도합니다.

### Q5. "경로가 너무 깁니다(MAX_PATH 260자)" 경고가 나옵니다.
> **A:** Windows 기본 파일 시스템 제한(260자)으로 인해 한글 COM 엔진이 긴 경로의 파일을 열지 못할 수 있습니다. HwpMate의 사전 점검(Preflight) 기능이 이를 미리 감지하여 변환 실패를 방지합니다. 폴더 경로를 `C:\문서`와 같이 상위 드라이브 쪽으로 이동하여 실행하시는 것을 권장합니다.

### Q6. 변환 중 한글 프로그램을 만져도 되나요?
> **A:** 변환 작업이 진행되는 동안에는 백그라운드에서 한글 프로세스가 초고속으로 문서를 열고 닫으므로, **한글 창을 직접 클릭하거나 조작하지 마세요.** 원활한 변환을 위해 변환 시작 전 열려있는 다른 한글 문서를 모두 저장하고 닫아주시는 것이 좋습니다.

---

## <a id="developer-guide"></a>🛠️ 개발자 및 기여 가이드

파이썬 개발 환경에서 직접 실행하거나 코드를 수정하고 빌드하려는 경우 아래 지침을 따릅니다.

### 1. 환경 설정 및 의존성 설치
```bash
git clone https://github.com/twbeatles/HwpMate.git
cd HwpMate

# 개발용 의존성 설치 (PyQt6, pywin32, pycryptodome, pytest 등)
pip install -r requirements-dev.txt
```

### 2. 소스코드 직접 실행
```powershell
# 관리자 권한 PowerShell에서 실행
python hwptopdf-hwpx_v4.py
```

### 3. 단일 실행 파일(EXE) 빌드
```powershell
pyinstaller --noconfirm --clean hwp_converter.spec
```
빌드가 완료되면 `dist/` 폴더 내에 단일 실행 파일이 생성됩니다.

### 4. 품질 검증 및 테스트 실행
HwpMate는 엔터프라이즈급 신뢰성을 위해 160개 이상의 단위/통합 테스트와 엄격한 타입 체킹을 유지합니다.
```powershell
# 정적 타입 검사
pyright .

# 전체 테스트 슈트 실행
pytest
```

---

## <a id="related-projects"></a>🔗 관련 프로젝트

* **단순·대량 HWP/HWPX 변환 (Batch Converter):** [HwpMate](https://github.com/twbeatles/HwpMate) (현재 저장소)
* **HWP 서식 편집, 매크로, 템플릿 자동화, 문서 관리:** [HwpMaster](https://github.com/twbeatles/HwpMaster)

---

## 🏷️ 검색 키워드 및 GitHub Topics
`hwp-to-pdf` `hwpx-to-pdf` `hwp-converter` `hwpx-converter` `batch-converter` `hwp-to-docx` `hancom-hangul` `hwp-to-image` `hwp-automation` `windows-batch` `한글-pdf-변환` `hwp-대량변환` `한글-word-변환`

