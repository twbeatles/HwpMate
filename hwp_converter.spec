# -*- mode: python ; coding: utf-8 -*-
"""
HWP 변환기 v8.7 - PyInstaller 빌드 설정
루트 래퍼 엔트리포인트(hwptopdf-hwpx_v4.py) 기준 경량화 빌드 설정
실제 애플리케이션 로직은 hwpmate/ 패키지에서 정적으로 import 됩니다.
2026-03-18 안정화/UX 보강(사전 점검, 결과 리포트, 안전한 강제 종료) 이후에도
추가 data 번들 없이 동일 빌드 구성이 동작함을 확인했습니다.
"""

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
    'PyQt6.QtCharts', 'PyQt6.QtStateMachine',
    'PyQt6.QtWebSockets', 'PyQt6.QtSerialBus',
    'PyQt6.QtSpatialAudio',
    
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
    datas=[],  # 문서/설정/리포트 템플릿은 런타임 생성하므로 배포 번들에 포함하지 않음
    hiddenimports=[
        # 필수 pywin32 모듈
        # 2026-03-18 기준 추가 hidden import 없이 빌드 검증 통과
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
        'qt6charts', 'qt6statemachine', 'qt6websockets',
        'qt6serialbus', 'qt6spatialaudio',
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
    name='HWP변환기_v8.7',
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
