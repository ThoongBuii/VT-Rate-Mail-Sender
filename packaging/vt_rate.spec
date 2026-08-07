# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec — build trên đúng OS cần phát hành.

import sys
from pathlib import Path

block_cipher = None
root = Path(SPECPATH).resolve()

datas = [
    (str(root / "app" / "web"), "app/web"),
    (str(root / "config.example.json"), "."),
    (str(root / "samples"), "samples"),
    (str(root / "templates"), "templates"),
]

hiddenimports = [
    "flask",
    "webview",
    "openpyxl",
    "app.webapp",
    "app.importer",
    "app.outlook_sender",
    "app.queue_worker",
    "app.template_engine",
    "app.models",
    "app.paths",
]

if sys.platform == "win32":
    hiddenimports += ["win32com", "win32com.client", "pythoncom", "pywintypes"]

a = Analysis(
    [str(root / "run.py")],
    pathex=[str(root)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["customtkinter", "tkinterweb"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="VTRateMailSender",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # cửa sổ app, không console
    disable_windowed_traceback=False,
    argv_emulation=False,
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
    upx=True,
    upx_exclude=[],
    name="VTRateMailSender",
)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="VT Rate Mail Sender.app",
        icon=None,
        bundle_identifier="vn.vtlogistics.ratemailsender",
        info_plist={
            "CFBundleName": "VT Rate Mail Sender",
            "CFBundleDisplayName": "VT Rate Mail Sender",
            "NSHighResolutionCapable": True,
            "NSAppleEventsUsageDescription": "Điều khiển Microsoft Outlook để gửi mail báo giá.",
        },
    )
