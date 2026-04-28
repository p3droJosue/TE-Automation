# build.spec — PyInstaller spec for TE Template Filler
# Usage (from the modular_architecture/ directory):
#   pip install pyinstaller customtkinter xlwings pandas
#   pyinstaller build.spec
#
# Output: dist/TE_Filler.exe  (single-file, no console window)

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# ── Paths ─────────────────────────────────────────────────────────────────────
HERE = Path(SPECPATH)   # directory containing this .spec file

# ── Data files ────────────────────────────────────────────────────────────────
datas = []

# CustomTkinter bundles its own theme JSON/image assets — they MUST be included
datas += collect_data_files("customtkinter")

# Optional: app icon (place a .ico file in assets/)
icon_path = str(HERE / "assets" / "icon.ico")
icon_arg  = icon_path if Path(icon_path).exists() else None

# ── Hidden imports ────────────────────────────────────────────────────────────
# xlwings uses COM on Windows; pandas has optional backends — collect them all
hiddenimports = []
hiddenimports += collect_submodules("xlwings")
hiddenimports += collect_submodules("pandas")
hiddenimports += ["customtkinter", "PIL", "PIL._imagingtk"]

# ── Analysis ──────────────────────────────────────────────────────────────────
a = Analysis(
    [str(HERE / "app.py")],
    pathex=[str(HERE)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "scipy", "notebook", "IPython"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="TE_Filler",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # no black terminal window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_arg,
    onefile=True,           # single .exe
)
