@echo off
REM Script to launch Word with VietNam Document Editor Add-in
REM Soạn thảo văn bản Việt Nam AI
REM Requires: npx serve installed globally

cd /d "%~dp0"

REM Start web server on port 3000
echo Starting web server on port 3000...
start "Web Server" cmd /k "npx serve -l 3000 -s"

REM Wait for server to start
timeout /t 3 /nobreak

REM Open Word with add-in from web server (using HTTP, not HTTPS)
echo Opening Word with add-in...
start winword.exe "http://localhost:3000/manifest.xml"

REM Exit this script
exit /b 0
