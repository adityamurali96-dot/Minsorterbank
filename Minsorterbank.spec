# PyInstaller spec for Minsorterbank
# Build with:  pyinstaller --clean -y Minsorterbank.spec
#
# Output:
#   Windows / Linux : dist/Minsorterbank.exe  (or dist/Minsorterbank)
#                     A single self-contained executable. Double-click to run.
#   macOS           : dist/Minsorterbank.app
#                     A single app bundle. Drag into /Applications.

# -*- mode: python ; coding: utf-8 -*-

import sys

from PyInstaller.utils.hooks import collect_submodules

IS_MAC = sys.platform == "darwin"

block_cipher = None

hiddenimports = []
hiddenimports += collect_submodules("pandas")
hiddenimports += collect_submodules("openpyxl")
hiddenimports += ["xlrd"]

datas = [
    ("app/templates", "app/templates"),
    ("app/static", "app/static"),
    ("sort_statement.py", "."),
]

a = Analysis(
    ["app/app.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "scipy", "PIL", "PyQt5", "PyQt6", "PySide6"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

if IS_MAC:
    # macOS: build a folder bundle and wrap it in a single .app.
    # The user sees one icon and drags it into /Applications.
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name="Minsorterbank",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=False,
        disable_windowed_traceback=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
    )

    coll = COLLECT(
        exe,
        a.binaries,
        a.zipfiles,
        a.datas,
        strip=False,
        upx=False,
        upx_exclude=[],
        name="Minsorterbank",
    )

    app = BUNDLE(
        coll,
        name="Minsorterbank.app",
        icon=None,
        bundle_identifier="com.minsorterbank.app",
        info_plist={
            "CFBundleName": "Minsorterbank",
            "CFBundleDisplayName": "Minsorterbank",
            "CFBundleShortVersionString": "1.0.0",
            "CFBundleVersion": "1.0.0",
            "NSHighResolutionCapable": True,
            "LSMinimumSystemVersion": "10.15.0",
        },
    )
else:
    # Windows / Linux: produce a single self-contained executable
    # (PyInstaller "onefile" mode). All Python, libs, templates and
    # static assets are embedded; the user gets exactly one file.
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.zipfiles,
        a.datas,
        [],
        name="Minsorterbank",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        runtime_tmpdir=None,
        console=False,
        disable_windowed_traceback=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
    )
