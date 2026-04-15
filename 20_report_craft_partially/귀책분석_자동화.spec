# -*- mode: python ; coding: utf-8 -*-
"""
귀책분석 자동화 시스템 — PyInstaller 빌드 스크립트

빌드 방법:
    pip install pyinstaller
    cd 20_report_craft_partially
    pyinstaller 귀책분석_자동화.spec

배포 폴더 구조 (dist/귀책분석_자동화/ 기준):
    귀책분석_자동화.exe        ← 실행 파일
    귀책분석_패턴집.md         ← exe 옆에 함께 복사 (빌드 후 자동 포함)
    reference/                 ← 참고 보고서 폴더 (사용자가 준비)
    output/                    ← 결과물 저장 폴더 (자동 생성)
    Tesseract-OCR/             ← OCR 엔진 (선택, 없으면 시스템 설치 사용)

주의:
    - reference/ 폴더는 사용자가 직접 배치해야 함
    - Tesseract-OCR/는 선택 사항 (없으면 C:\\Program Files\\Tesseract-OCR\\ 자동 탐색)
"""

import sys
from pathlib import Path

block_cipher = None

# ── 분석 ─────────────────────────────────────────────────────────────────────
a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # 귀책분석 패턴집 — BASE_DIR 바로 아래에 위치해야 하므로 '.' 로 지정
        ('귀책분석_패턴집.md', '.'),
        # 레퍼런스 학습 결과 — 배포 전 미리 생성해서 포함
        # (output/reference_patterns.md 가 존재하면 포함, 없으면 빌드 전 learn 실행)
        ('output/reference_patterns.md', 'output'),
    ],
    hiddenimports=[
        # PyMuPDF (fitz)
        'fitz',
        'fitz.fitz',
        # pdfplumber
        'pdfplumber',
        'pdfminer',
        'pdfminer.high_level',
        'pdfminer.layout',
        'pdfminer.converter',
        'pdfminer.pdfinterp',
        'pdfminer.pdfdevice',
        # python-docx
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        # pytesseract / Pillow
        'pytesseract',
        'PIL',
        'PIL.Image',
        # tqdm
        'tqdm',
        # python-dotenv
        'dotenv',
        # src 모듈
        'src',
        'src.text_extractor',
        'src.correspondence_scanner',
        'src.prompt_builder',
        'src.report_generator',
        'src.reference_learner',
        # 표준 라이브러리 (간혹 누락)
        'xml.etree.ElementTree',
        'zipfile',
        'shutil',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Claude API 관련 — 사용 안 함
        'anthropic',
        # 불필요한 대형 패키지
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'sklearn',
        'tensorflow',
        'torch',
        'jupyter',
        'IPython',
        'notebook',
        'tkinter',
        'wx',
        'PyQt5',
        'PyQt6',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# ── exe 생성 ─────────────────────────────────────────────────────────────────
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='귀책분석_자동화',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,          # CLI 도구 — 콘솔 창 유지
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',     # 아이콘 파일이 있으면 주석 해제
)

# ── 폴더 배포 (--onedir 모드) ─────────────────────────────────────────────────
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='귀책분석_자동화',
)
