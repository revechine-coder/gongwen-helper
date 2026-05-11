$ErrorActionPreference = "Stop"

$RootDir = Resolve-Path (Join-Path $PSScriptRoot "../..")
Set-Location $RootDir
$env:PYINSTALLER_CONFIG_DIR = Join-Path $RootDir ".pyinstaller-cache"

if (-not (Test-Path ".venv-packaging\Scripts\python.exe")) {
  $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
  if ($pythonCmd) {
    python -m venv .venv-packaging
  } else {
    $pythonLauncher = Get-Command py -ErrorAction SilentlyContinue
    if (-not $pythonLauncher) {
      throw "未找到 Python。请先安装 Python 3.11 或更高版本。"
    }
    py -3.11 -m venv .venv-packaging
  }
}

$PythonExe = ".\.venv-packaging\Scripts\python.exe"
$Version = & $PythonExe -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"
if ([version]$Version -lt [version]"3.11") {
  throw "Python 版本过低：$Version。请使用 Python 3.11 或更高版本。"
}

.\.venv-packaging\Scripts\python.exe -m pip install -r requirements-packaging.txt
.\.venv-packaging\Scripts\python.exe -m PyInstaller `
  --clean `
  --noconfirm `
  --windowed `
  --onefile `
  --name "公文格式助手" `
  --add-data "public;public" `
  --add-data "storage;storage" `
  desktop_app.py

Copy-Item -Recurse -Force "public" "dist\public"

Write-Host "Windows exe generated at: $RootDir\dist\公文格式助手.exe"
Write-Host "Web assets copied to: $RootDir\dist\public"
