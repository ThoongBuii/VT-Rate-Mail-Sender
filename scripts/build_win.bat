@echo off
REM Build Windows onefile .exe
cd /d %~dp0\..
python -m venv .venv
call .venv\Scripts\activate
pip install -r requirements.txt -r requirements-build.txt
pip install pywin32
rmdir /s /q build dist 2>nul
pyinstaller packaging\vt_rate.spec --noconfirm
echo.
echo Xong: dist\VTRateMailSender.exe  (1 file — copy di dau cung duoc)
pause
