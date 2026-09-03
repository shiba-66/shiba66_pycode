# est_genka.py / dnd_filepicker.py のセットアップ用インストーラー。
# Python確認/インストール -> 必要ライブラリのインストール -> コード取得 -> デスクトップショートカット作成

$ErrorActionPreference = "Stop"

$RepoOwner = "shiba-66"
$RepoName = "shiba66_pycode"
$Branch = "main"
$InstallDir = "C:\pycode"
$Files = @("est_genka.py", "dnd_filepicker.py")
$Requirements = @("pandas", "openpyxl", "pyodbc", "tkinterdnd2")
$PythonWingetId = "Python.Python.3.13"

function Write-Step($msg) {
    Write-Host "==> $msg" -ForegroundColor Cyan
}

# ---------------------------------------------------------------------------
# 1. Pythonの確認・インストール
# ---------------------------------------------------------------------------
Write-Step "Pythonの確認..."
$python = Get-Command python -ErrorAction SilentlyContinue

if (-not $python) {
    Write-Step "Pythonが見つからないため winget でインストールします"
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        Write-Error "wingetが見つかりません。手動でPythonをインストールしてから再実行してください: https://www.python.org/downloads/"
        exit 1
    }
    winget install -e --id $PythonWingetId --source winget `
        --accept-package-agreements --accept-source-agreements --silent

    # インストーラーが更新したPATHをこのプロセスに反映
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + `
                [System.Environment]::GetEnvironmentVariable("Path", "User")

    $python = Get-Command python -ErrorAction SilentlyContinue
    if (-not $python) {
        Write-Error "Pythonのインストールを確認できませんでした。PCを再起動してから、このインストーラーを再実行してください。"
        exit 1
    }
} else {
    Write-Step "Pythonは既にインストール済みです: $($python.Source)"
}

# ---------------------------------------------------------------------------
# 2. 必要ライブラリのインストール
# ---------------------------------------------------------------------------
Write-Step "必要なライブラリをインストールします: $($Requirements -join ', ')"
python -m pip install --upgrade pip | Out-Null
python -m pip install @Requirements
if ($LASTEXITCODE -ne 0) {
    Write-Error "ライブラリのインストールに失敗しました。"
    exit 1
}

# ---------------------------------------------------------------------------
# 3. Microsoft Access Driver の確認 (見積物件管理DBへの接続に必須)
# ---------------------------------------------------------------------------
Write-Step "Microsoft Access Driverの確認..."
$hasAccessDriver = $false
try {
    $hasAccessDriver = [bool](Get-OdbcDriver -Name "*Access*" -ErrorAction SilentlyContinue)
} catch {}

if (-not $hasAccessDriver) {
    Write-Warning "『Microsoft Access Driver (*.mdb, *.accdb)』が見つかりませんでした。"
    Write-Warning "est_genka.py のDB連携(仕分ワード検索)が失敗する可能性があります。"
    Write-Warning "Office/Accessが入っていないPCでは、別途「Microsoft Access Database Engine 再頒布可能パッケージ」の導入が必要です。"
}

# ---------------------------------------------------------------------------
# 4. コードのダウンロード
# ---------------------------------------------------------------------------
Write-Step "$InstallDir を作成します"
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null

foreach ($file in $Files) {
    $url = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch/$file"
    $dest = Join-Path $InstallDir $file
    Write-Step "$file をダウンロード中..."
    Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
}

# ---------------------------------------------------------------------------
# 5. デスクトップショートカット作成
# ---------------------------------------------------------------------------
Write-Step "デスクトップショートカットを作成します"

$pythonDir = Split-Path -Parent $python.Source
$pythonw = Join-Path $pythonDir "pythonw.exe"
if (-not (Test-Path $pythonw)) {
    $pythonw = $python.Source
}

$WshShell = New-Object -ComObject WScript.Shell
$ShortcutPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "見積仕分け(est_genka).lnk"
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $pythonw
$Shortcut.Arguments = '"' + (Join-Path $InstallDir "est_genka.py") + '"'
$Shortcut.WorkingDirectory = $InstallDir
$Shortcut.IconLocation = "$pythonw,0"
$Shortcut.Description = "見積書xlsxをドラッグ&ドロップすると仕分け済みxlsxを作成します"
$Shortcut.Save()

Write-Step "完了しました。デスクトップの『見積仕分け(est_genka)』アイコンに見積書xlsxをドラッグ&ドロップしてください。"
