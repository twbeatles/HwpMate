# HWP/HWPX 변환기 v8.0 (PyQt6)

한글(HWP/HWPX) 파일을 PDF, HWPX, 또는 **DOCX**로 일괄 변환하는 Windows용 GUI 프로그램

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![PyQt6](https://img.shields.io/badge/PyQt6-6.0+-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)

## ✨ 주요 기능

- **📕 PDF / 📘 HWPX / 📄 DOCX** 변환 지원
- **폴더 일괄 변환** 또는 **파일 개별 선택**
- **드래그 앤 드롭** 파일 추가
- **다크/라이트 테마**
- **Toast 알림 시스템**
- **HiDPI 지원**
- **예상 남은 시간 표시**

## 💻 요구사항

- Windows 10/11
- Python 3.9+
- **한컴오피스 한글 2018 이상** (필수)
- 관리자 권한

## 📦 설치 및 실행

```bash
pip install pywin32 PyQt6
python hwptopdf-hwpx_v4.py  # 관리자 권한 필요
```

### 빌드
```bash
pyinstaller hwp_converter.spec
```

## 🛠️ v8.0 신규

- ✅ DOCX 변환 (OOXML 포맷)
- ✅ Toast 알림
- ✅ HiDPI 지원  
- ✅ 예상 시간/소요 시간 표시

---
⭐ 도움이 되었다면 Star를 눌러주세요!
