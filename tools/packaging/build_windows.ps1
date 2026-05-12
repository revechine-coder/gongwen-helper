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
      throw "Python not found. Please install Python 3.11 or higher."
    }
    py -3.11 -m venv .venv-packaging
  }
}

$PythonExe = ".\.venv-packaging\Scripts\python.exe"
$Version = & $PythonExe -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"
if ([version]$Version -lt [version]"3.11") {
  throw "Python version too low: $Version. Please use Python 3.11 or higher."
}

.\.venv-packaging\Scripts\python.exe -m pip install -r requirements-packaging.txt
if (Test-Path ".packaging-storage") {
  Remove-Item -Recurse -Force ".packaging-storage"
}
New-Item -ItemType Directory -Force ".packaging-storage\uploads" | Out-Null
New-Item -ItemType Directory -Force ".packaging-storage\outputs" | Out-Null
.\.venv-packaging\Scripts\python.exe -m PyInstaller `
  --clean `
  --noconfirm `
  --windowed `
  --onefile `
  --name "GongwenHelper" `
  --add-data "public;public" `
  --add-data ".packaging-storage;storage" `
  desktop_app.py

Copy-Item -Recurse -Force "public" "dist\public"

Write-Host "Windows exe generated at: $RootDir\dist\GongwenHelper.exe"
Write-Host "Web assets copied to: $RootDir\dist\public"
