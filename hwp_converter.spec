# -*- mode: python ; coding: utf-8 -*-
"""
HWP 변환기 v8.4 - PyInstaller 빌드 설정
경량화 빌드를 위한 최적화 설정 적용
"""

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# 제외할 불필요한 모듈 목록 (경량화)
EXCLUDES = [
    # 테스트/디버깅
    'pytest', 'unittest', 'test', 'tests',
    
    # 사용하지 않는 PyQt6 모듈
    'PyQt6.QtWebEngine', 'PyQt6.QtWebEngineCore', 'PyQt6.QtWebEngineWidgets',
    'PyQt6.QtMultimedia', 'PyQt6.QtMultimediaWidgets',
    'PyQt6.QtBluetooth', 'PyQt6.QtNfc',
    'PyQt6.QtQuick', 'PyQt6.QtQuick3D', 'PyQt6.QtQml',
    'PyQt6.QtSql', 'PyQt6.QtNetwork',
    'PyQt6.QtOpenGL', 'PyQt6.QtOpenGLWidgets',
    'PyQt6.QtSvg', 'PyQt6.QtSvgWidgets',
    'PyQt6.QtPdf', 'PyQt6.QtPdfWidgets',
    'PyQt6.QtDesigner', 'PyQt6.QtHelp',
    'PyQt6.QtRemoteObjects', 'PyQt6.QtSensors',
    'PyQt6.QtSerialPort', 'PyQt6.QtPositioning',
    'PyQt6.QtTextToSpeech', 'PyQt6.Qt3DCore',
    'PyQt6.Qt3DInput', 'PyQt6.Qt3DLogic',
    'PyQt6.Qt3DRender', 'PyQt6.Qt3DExtras',
    
    # 사용하지 않는 기타 모듈
    'PIL', 'numpy', 'pandas', 'matplotlib',
    'scipy', 'sklearn', 'tensorflow', 'torch',
    'tkinter', 'tk', 'tcl',
    'IPython', 'jupyter',
    'cryptography', 'ssl',
    'asyncio', 'concurrent',
    'xml.etree', 'html.parser',
    'lib2to3', 'distutils',
]

a = Analysis(
    ['hwptopdf-hwpx_v4.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        # 필수 pywin32 모듈
        'win32com.client',
        'win32api',
        'pythoncom',
        'pywintypes',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 불필요한 바이너리 제거 (경량화)
a.binaries = [
    b for b in a.binaries 
    if not any(x in b[0].lower() for x in [
        'qt6webengine', 'qt6multimedia', 'qt6quick',
        'qt6qml', 'qt6sql', 'qt6network', 'qt6opengl',
        'qt6svg', 'qt6pdf', 'qt6designer',
        'd3dcompiler', 'opengl32sw',
    ])
]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='HWP변환기_v8.4',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,  # Windows에서는 strip 효과 제한적
    upx=True,  # UPX 압축 활성화 (경량화)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI 모드 (콘솔 창 숨김)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 아이콘이 있으면 경로 지정: 'icon.ico'
    uac_admin=True,  # 관리자 권한 요청
)
