# PyInstaller 빌드 설정: macOS 는 Zoteley.app, Windows 는 Zoteley 폴더(Zoteley.exe)
# 실행: pyinstaller --noconfirm zoteley.spec
import sys
from pathlib import Path

ROOT = Path(SPECPATH)
ASSETS = ROOT / "assets"
VERSION = "1.1.1"
datas = [(str(p), "assets") for p in ASSETS.glob("*.png")] if ASSETS.exists() else []
icon_mac = ASSETS / "icon.icns"
icon_win = ASSETS / "icon.ico"

a = Analysis(
    ["gui.py"],
    pathex=[str(ROOT)],
    datas=datas,
    hiddenimports=["zms"],
    # 쓰지 않는 Qt 모듈을 빼서 설치 파일 크기를 줄인다
    excludes=["PySide6.QtNetwork", "PySide6.QtQml", "PySide6.QtQuick", "PySide6.QtSql",
              "PySide6.QtTest", "PySide6.QtXml", "tkinter"],
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="Zoteley",
    console=False,
    icon=str(icon_win) if sys.platform == "win32" and icon_win.exists() else None,
)
coll = COLLECT(exe, a.binaries, a.datas, name="Zoteley")

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="Zoteley.app",
        icon=str(icon_mac) if icon_mac.exists() else None,
        bundle_identifier="io.github.hspark-rm.zoteley",
        version=VERSION,
        info_plist={
            "CFBundleShortVersionString": VERSION,
            "NSHighResolutionCapable": True,
            # Dock 아이콘에 Word 원고를 끌어다 놓을 수 있게 한다
            "CFBundleDocumentTypes": [{
                "CFBundleTypeName": "Word Document",
                "CFBundleTypeRole": "Editor",
                "LSHandlerRank": "Alternate",
                "LSItemContentTypes": ["org.openxmlformats.wordprocessingml.document"],
            }],
        },
    )
