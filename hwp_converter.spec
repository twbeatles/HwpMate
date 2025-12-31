# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for HWP/HWPX 변환기 v8.3

빌드 명령:
    pyinstaller hwp_converter.spec

출력:
    dist/HWP변환기_v8.3.exe

변경 이력:
    v8.3 - 드래그 앤 드롭 수정 (관리자 권한 호환)
    v8.2 - 디버깅 및 리팩토링
    v8.1 - UI/UX 개선
    v8.0 - DOCX 변환 지원
"""

import sys
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

# =============================================================================
# pywin32 관련 모듈 수집
# =============================================================================
pywin32_datas = []
pywin32_binaries = []
pywin32_hiddenimports = []

# win32com 패키지 전체 수집
try:
    win32com_datas, win32com_binaries, win32com_hiddenimports = collect_all('win32com')
    pywin32_datas.extend(win32com_datas)
    pywin32_binaries.extend(win32com_binaries)
    pywin32_hiddenimports.extend(win32com_hiddenimports)
except Exception:
    pass

# pythoncom 수집
try:
    pythoncom_datas, pythoncom_binaries, pythoncom_hiddenimports = collect_all('pythoncom')
    pywin32_datas.extend(pythoncom_datas)
    pywin32_binaries.extend(pythoncom_binaries)
    pywin32_hiddenimports.extend(pythoncom_hiddenimports)
except Exception:
    pass

# =============================================================================
# Analysis 설정
# =============================================================================
a = Analysis(
    ['hwptopdf-hwpx_v4.py'],
    pathex=[],
    binaries=pywin32_binaries,
    datas=pywin32_datas,
    hiddenimports=[
        # pywin32 핵심 모듈
        'win32com',
        'win32com.client',
        'win32com.client.gencache',
        'win32com.client.dynamic',
        'win32com.gen_py',
        'pythoncom',
        'pywintypes',
        'win32api',
        'win32con',
        # PyQt6 모듈
        'PyQt6',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
        'PyQt6.QtWidgets',
        'PyQt6.sip',
    ] + pywin32_hiddenimports + collect_submodules('win32com'),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 사용하지 않는 모듈 제외 (빌드 크기 최적화)
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'PIL',
        'cv2',
        'IPython',
        'notebook',
        'pytest',
        'unittest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# =============================================================================
# PYZ 아카이브
# =============================================================================
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# =============================================================================
# EXE 설정
# =============================================================================
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='HWP변환기_v8.3',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # UPX 압축 사용 (빌드 크기 감소)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI 앱이므로 콘솔 숨김
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 아이콘이 있다면 경로 지정: icon='icon.ico'
    version_info=None,
    uac_admin=True,  # 관리자 권한 요청 (한글 COM 접근에 필요)
)
