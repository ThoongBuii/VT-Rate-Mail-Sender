@echo off
REM Build Windows .exe (chạy trên máy Windows)
cd /d %~dp0\..
python -m venv .venv
call .venv\Scripts\activate
pip install -r requirements.txt -r requirements-build.txt
pip install pywin32
rmdir /s /q build dist 2>nul
pyinstaller packaging\vt_rate.spec --noconfirm
echo.
echo Xong: dist\VTRateMailSender\
echo Zip thu muc do de phat hanh van phong.
pause
