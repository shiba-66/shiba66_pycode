@echo off
REM est_genka.py セットアップ用インストーラー(入口)
REM このファイルだけ保存して実行すれば、本体スクリプトをGitHubから取得して実行します。

setlocal
set REPO_RAW=https://raw.githubusercontent.com/shiba-66/shiba66_pycode/main
set SCRIPT=%TEMP%\install_est_genka.ps1

echo インストーラー本体をダウンロードしています...
powershell -NoProfile -Command "Invoke-WebRequest -Uri '%REPO_RAW%/install_est_genka.ps1' -OutFile '%SCRIPT%' -UseBasicParsing"

if not exist "%SCRIPT%" (
    echo ダウンロードに失敗しました。ネットワーク接続を確認してください。
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
pause
