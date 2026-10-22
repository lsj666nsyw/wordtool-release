# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['wordtool.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['docx', 'PIL', 'lxml'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pandas', 'matplotlib', 'scipy', 'pymupdf', 'fitz', 'pypdf', 'pypdfium2', 'reportlab'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)


exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Word批量处理助手',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Word批量处理助手',
)
