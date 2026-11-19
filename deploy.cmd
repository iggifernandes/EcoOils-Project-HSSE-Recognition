@echo off
rem Push Code.gs to Apps Script and update the LIVE web-app deployment (same URL, new version).
rem Windows version of deploy.sh. Run from the repo folder:  deploy.cmd
cd /d "%~dp0"
set DEPLOYMENT_ID=AKfycbx2wdQ_TGsu6Kfj8xk8Tjz8vuqvzdKTuKKlLZFb214dai5J32Kw781Z1DOZ9Amd3GNV
call clasp push -f || exit /b 1
call clasp deploy --deploymentId %DEPLOYMENT_ID% --description "clasp %DATE% %TIME%" || exit /b 1
echo.
echo Done. Check the version at https://script.google.com/macros/s/%DEPLOYMENT_ID%/exec
