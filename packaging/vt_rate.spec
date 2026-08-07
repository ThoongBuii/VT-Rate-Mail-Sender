# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec — Windows: 1 file .exe · macOS: .app bundle

import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_all, collect_dynamic_libs

block_cipher = None
root = Path(SPECPATH).resolve().parent
icon_ico = root / "packaging" / "icons" / "app.ico"
icon_icns = root / "packaging" / "icons" / "app.icns"

datas = [
    (str(root / "app" / "web"), "app/web"),
    (str(root / "config.example.json"), "."),
    (str(root / "samples"), "samples"),
    (str(root / "templates"), "templates"),
]
binaries = []
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
    "app.clipboard_html",
]

# Gom toàn bộ resource pywebview (WebView2 loader, JS, …)
wv_datas, wv_binaries, wv_hidden = collect_all("webview")
datas += wv_datas
binaries += wv_binaries
hiddenimports += wv_hidden
binaries += collect_dynamic_libs("webview")

if sys.platform == "win32":
    hiddenimports += [
        "win32com",
        "win32com.client",
        "pythoncom",
        "pywintypes",
        "webview.platforms.edgechromium",
        "clr_loader",
    ]

a = Analysis(
    [str(root / "run.py")],
    pathex=[str(root)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[str(root / "packaging" / "pyi_rth_vt_path.py")],
    excludes=["customtkinter", "tkinterweb"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

if sys.platform == "win32":
    # Một file .exe độc lập — có thể copy đi đâu cũng chạy (cần WebView2 Runtime hệ thống)
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.zipfiles,
        a.datas,
        [],
        name="VTRateMailSender",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=str(icon_ico) if icon_ico.exists() else None,
    )
else:
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
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=str(icon_ico) if icon_ico.exists() else None,
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

    app = BUNDLE(
        coll,
        name="VT Rate Mail Sender.app",
        icon=str(icon_icns) if icon_icns.exists() else None,
        bundle_identifier="vn.vtlogistics.ratemailsender",
        info_plist={
            "CFBundleName": "VT Rate Mail Sender",
            "CFBundleDisplayName": "VT Rate Mail Sender",
            "CFBundleShortVersionString": "1.0.0",
            "CFBundleVersion": "1.0.0",
            "NSHighResolutionCapable": True,
            "NSAppleEventsUsageDescription": "Điều khiển Microsoft Outlook để gửi mail báo giá.",
        },
    )
