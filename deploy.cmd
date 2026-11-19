@echo off
rem Double-click this file. It fetches the latest code from GitHub, pushes Code.gs to Apps Script,
rem and updates the LIVE web app (same URL, new version).
cd /d "%~dp0"
set DEPLOYMENT_ID=AKfycbx2wdQ_TGsu6Kfj8xk8Tjz8vuqvzdKTuKKlLZFb214dai5J32Kw781Z1DOZ9Amd3GNV
echo Getting the latest code from GitHub...
git pull || goto :fail
echo.
echo Sending Code.gs to Apps Script...
call clasp push -f || goto :fail
call clasp deploy --deploymentId %DEPLOYMENT_ID% --description "clasp %DATE% %TIME%" || goto :fail
echo.
echo ===== DONE. The live backend is updated. You can close this window. =====
pause
exit /b 0
:fail
echo.
echo ===== SOMETHING WENT WRONG. Take a photo or screenshot of this window and send it to Claude. =====
pause
exit /b 1
