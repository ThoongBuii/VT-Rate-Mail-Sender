@echo off
REM Tat VT Rate Mail Sender + Outlook dang treo (Windows)
REM Chay file nay khi khong xoa duoc thu muc / Outlook khong mo

echo Dang tat VTRateMailSender.exe (neu co)...
taskkill /F /IM VTRateMailSender.exe >nul 2>&1

echo Dang tat OUTLOOK.EXE (neu treo)...
taskkill /F /IM OUTLOOK.EXE >nul 2>&1

REM Mot so may con process Office
taskkill /F /IM OfficeClickToRun.exe >nul 2>&1

echo.
echo Xong. Hay:
echo  1) Thu xoa / ghi de thu muc VTRateMailSender
echo  2) Mo lai Microsoft Outlook
echo.
pause
