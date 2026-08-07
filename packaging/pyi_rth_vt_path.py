# PyInstaller runtime hook — thêm _MEIPASS vào PATH (WebView2Loader.dll, …)
import os
import sys

if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    meipass = sys._MEIPASS
    os.environ["PATH"] = meipass + os.pathsep + os.environ.get("PATH", "")
    # pywebview / Edge WebView2 đôi khi tìm DLL theo cwd
    try:
        os.chdir(meipass)
    except OSError:
        pass
