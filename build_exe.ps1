param(
    [string]$AppName = "pptx_to_text"
)

$ErrorActionPreference = "Stop"

if (!(Test-Path ".\venv\Scripts\python.exe")) {
    throw "Virtual environment python not found at .\venv\Scripts\python.exe"
}

Write-Host "Installing/updating build dependencies..."
.\venv\Scripts\python.exe -m pip install --upgrade pip pyinstaller python-pptx

Write-Host "Building standalone executable..."
.\venv\Scripts\python.exe -m PyInstaller `
    --onefile `
    --noconsole `
    --name $AppName `
    --clean `
    .\pptx_to_text.py

Write-Host ""
Write-Host "Build complete:"
Write-Host "  .\dist\$AppName.exe"
