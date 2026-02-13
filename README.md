# HwpMate 

한컴오피스 한글(HWP/HWPX) 파일을 **PDF, HWPX, DOCX, HTML, ODT, RTF, TXT, 이미지(PNG, JPG, BMP, GIF)** 등 다양한 포맷으로 손쉽게 일괄 변환하는 Windows용 GUI 프로그램입니다.

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![PyQt6](https://img.shields.io/badge/PyQt6-6.0+-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows_10/11-lightgrey.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

---

## ✨ 주요 기능

### 변환 기능
| 형식 | 설명 |
|------|------|
| 📕 **PDF** | 문서 공유에 적합한 표준 형식 |
| 📘 **HWPX** | 한글 호환 XML 기반 형식 |
| 📄 **DOCX** | MS Word 호환 OOXML 형식 |
| 🌐 **HTML/ODT** | 웹 및 오픈 문서 형식 지원 |
| 🖼️ **이미지** | PNG, JPG, BMP, GIF 이미지 변환 지원 |

### HwpMate만의 특징
- **탭 기반 인터페이스** - '문서 변환'과 '이미지 변환'을 탭으로 구분하여 정리
- **변환 형식 카드 UI** - 시각적으로 직관적인 형식 선택 (텍스트 잘림 방지 개선)
- **자동 백업** - 변환 전 원본 파일을 안전하게 백업 (`backup` 폴더 자동 생성)
- **폴더 일괄 변환** - 폴더 내 모든 HWP/HWPX 파일 일괄 처리
- **파일 개별 선택** - 원하는 파일만 선택하여 변환
- **드래그 앤 드롭** - 파일 또는 폴더를 드래그하여 추가 (**관리자 권한에서도 완벽 동작**)
- **다크/라이트 테마** - 모던한 디자인, 메뉴바/상태바 스타일 적용
- **Toast 알림** - 스택 기능 지원 (최대 3개 동시 표시)
- **시스템 트레이** - 최소화 시 트레이로 숨김
- **테이블 호버 효과** - 행 선택 시 시각적 피드백
- **HiDPI 지원** - 고해상도 디스플레이 지원
- **예상 시간 표시** - 변환 남은 시간 실시간 계산

---

## 💻 시스템 요구사항

| 항목 | 요구사항 |
|------|----------|
| 운영체제 | Windows 10/11 (64-bit) |
| Python | 3.9 이상 |
| 한컴오피스 | 한글 2018 이상 (**필수**) |
| 권한 | **관리자 권한** 필요 |

---

## 📦 설치

### 1. 의존성 설치
```bash
pip install pywin32 PyQt6
```

### 2. 실행
```bash
# 반드시 관리자 권한으로 실행
python hwptopdf-hwpx_v4.py
```

### 3. 빌드 (선택사항)
```bash
# PyInstaller로 실행 파일 생성
pyinstaller hwp_converter.spec

# 생성된 파일 위치
# dist/HWP변환기_v8.4.exe
```

---

## ⌨️ 키보드 단축키

| 단축키 | 동작 |
|--------|------|
| `Ctrl+O` | 파일 추가 |
| `Ctrl+Shift+O` | 폴더 선택 |
| `Ctrl+Enter` | 변환 시작 |
| `Esc` | 변환 취소 |
| `Delete` | 선택 파일 제거 |
| `Ctrl+Delete` | 전체 파일 제거 |
| `F1` | 프로그램 정보 |

---

## 📁 프로젝트 구조

```
hwp-to-pdf-hwpx/
├── hwptopdf-hwpx_v4.py    # 메인 프로그램
├── hwp_converter.spec      # PyInstaller 빌드 설정 (경량화)
├── README.md               # 문서
└── update_history.md       # 업데이트 이력
```

---

## 🛠️ 버전 히스토리

### v8.6 (2026-01-26) - 포맷 확장 및 UI 개선
**기능 확장:**
- **다양한 포맷 지원**: PDF, HWPX, DOCX 외 ODT, HTML, RTF, TXT 및 이미지(PNG, JPG, BMP, GIF) 지원
- **자동 백업**: 변환 시작 전 원본 파일을 `backup` 폴더에 안전하게 복사

**UI/UX 개선:**
- **탭 인터페이스 도입**: 문서/이미지 변환 탭 분리
- **카드 디자인 개선**: 가독성 향상 및 레이아웃 최적화

### v8.5 (2026-01-15) - 안정성 및 UX 개선
**안정성 강화:**
- 변환 중 파일 목록 변경 방지 (버튼/드롭 영역 비활성화)
- 출력 폴더 쓰기 권한 사전 검사
- 파일 경로 유효성 검증 함수 추가
- 테이블 정렬 비활성화 (데이터 동기화 문제 방지)

**UX 개선:**
- 폴더 선택 시 HWP/HWPX 파일 수 미리보기
- 대용량 파일(50개+) 추가 시 진행 상태 표시
- 실패 목록 텍스트 파일 내보내기 기능
- 시스템 트레이 아이콘 표시 수정

### v8.4.1 (2026-01-06) - 코드 품질 개선
**UI/UX 개선:**
- **FormatCard 카드 UI** - 변환 형식 선택을 시각적 카드로 개선
- **테마 스타일 강화** - 메뉴바, 상태바, 테이블 호버 효과
- **진행률 바 글로우** - 시각적 개선

**코드 품질:**
- 상수 추출 (윈도우 크기, 타이머, 변환 대기 시간)
- 예외 처리 강화
- 조건부 로깅 최적화
- 리소스 관리 개선 (ToastWidget/Manager)
- 파일 중복 체크 O(1) 성능

### v8.4 (2026-01-01) - 네이티브 드래그 앤 드롭
- **네이티브 Windows 드래그 앤 드롭** - 관리자 권한에서도 100% 동작
- Windows Shell API (`WM_DROPFILES`) 직접 사용
- 64비트 핸들 오버플로 수정

### v8.1~8.3
- 메뉴바, 상태바, 시스템 트레이 추가
- 키보드 단축키, 툴팁 지원
- DOCX 변환 지원, HiDPI 지원

---

## ⚠️ 주의사항

1. **관리자 권한**: 한글 COM 객체 접근 및 드래그 앤 드롭을 위해 필요
2. **한글 설치**: 한컴오피스 한글이 설치되어 있어야 합니다
3. **변환 중 한글 사용 금지**: 변환 중에는 한글 프로그램을 직접 사용하지 마세요

---

## 🔧 기술 스택

- **GUI Framework**: PyQt6
- **COM Automation**: pywin32 (win32com)
- **Drag & Drop**: Windows Shell API (WM_DROPFILES)
- **Build Tool**: PyInstaller (경량화 빌드)

---

## 📄 라이선스

MIT License | Copyright (c) 2025-2026

---

⭐ **도움이 되었다면 Star를 눌러주세요!**
