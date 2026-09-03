@echo off
REM est_genka.py setup installer (entry point)
REM Save this file and run it. It fetches the real installer script from GitHub and runs it.
REM (This file is kept ASCII-only on purpose: cmd.exe reads .bat files using the
REM  system's legacy codepage, and non-ASCII text here caused garbled parsing on
REM  some PCs. The Japanese messages are shown by install_est_genka.ps1 instead.)

setlocal
set REPO_RAW=https://raw.githubusercontent.com/shiba-66/shiba66_pycode/main
set SCRIPT=%TEMP%\install_est_genka.ps1

echo Downloading installer script...
powershell -NoProfile -Command "Invoke-WebRequest -Uri '%REPO_RAW%/install_est_genka.ps1' -OutFile '%SCRIPT%' -UseBasicParsing"

if not exist "%SCRIPT%" (
    echo Download failed. Please check your network connection.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
pause
