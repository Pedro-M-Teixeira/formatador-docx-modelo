# -*- mode: python ; coding: utf-8 -*-
"""
Especificação do PyInstaller para o Formatador DOCX por Modelo.

Produz um executável único, sem consola, com o nome:
    Windows : FormatadorDocx.exe
    macOS   : FormatadorDocx  (+ FormatadorDocx.app)
    Linux   : FormatadorDocx

Compilar localmente:
    pip install -r requirements.txt pyinstaller
    pyinstaller build.spec --noconfirm --clean
O resultado fica em dist/.
"""

import sys
from pathlib import Path

RAIZ = Path(SPECPATH)  # noqa: F821  (SPECPATH é injectado pelo PyInstaller)
JANELA = sys.platform.startswith("win") or sys.platform == "darwin"

icone = None
if sys.platform.startswith("win"):
    icone = str(RAIZ / "assets" / "icone.ico")
elif sys.platform == "darwin":
    icns = RAIZ / "assets" / "icone.icns"
    icone = str(icns) if icns.exists() else None

ficheiro_versao = None
if sys.platform.startswith("win"):
    vi = RAIZ / "version_info.txt"
    ficheiro_versao = str(vi) if vi.exists() else None

a = Analysis(
    [str(RAIZ / "formatar_docx_modelo_gui.py")],
    pathex=[str(RAIZ)],
    binaries=[],
    datas=[],
    hiddenimports=[
        "lxml._elementpath",
        "lxml.etree",
        "tkinter",
        "tkinter.filedialog",
        "tkinter.messagebox",
        "tkinter.ttk",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "numpy", "PIL", "matplotlib", "pandas", "scipy", "pytest",
        "setuptools", "pip", "test", "unittest",
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)  # noqa: F821

exe = EXE(  # noqa: F821
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="FormatadorDocx",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                 # UPX desactivado: aumenta falsos positivos de antivírus
    runtime_tmpdir=None,
    console=not JANELA,        # sem janela de consola no Windows/macOS
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icone,
    version=ficheiro_versao,
)

if sys.platform == "darwin":
    app = BUNDLE(  # noqa: F821
        exe,
        name="FormatadorDocx.app",
        icon=icone,
        bundle_identifier="pt.formatadordocx.app",
        info_plist={
            "CFBundleDisplayName": "Formatador DOCX",
            "CFBundleShortVersionString": "1.0.0",
            "NSHighResolutionCapable": True,
        },
    )
