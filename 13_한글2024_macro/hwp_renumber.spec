# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec 파일
# 빌드 방법: pyinstaller hwp_renumber.spec

block_cipher = None

a = Analysis(
    ['hwp_renumber.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'win32com',
        'win32com.client',
        'pywintypes',
        'pythoncom',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='hwp_renumber',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,       # GUI 앱이므로 터미널 창 숨김
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,           # 아이콘 경로 지정 시: icon='icon.ico'
    onefile=True,        # 단일 exe 파일로 생성
)
