param(
    [string]$AppName = "pptx_to_text"
)

$ErrorActionPreference = "Stop"

if (!(Test-Path ".\venv\Scripts\python.exe")) {
    throw "Virtual environment python not found at .\venv\Scripts\python.exe"
}
if (!(Test-Path ".\assets\app_icon.ico")) {
    throw "Icon file not found at .\assets\app_icon.ico"
}

Write-Host "Installing/updating build dependencies..."
.\venv\Scripts\python.exe -m pip install --upgrade pip pyinstaller python-pptx

Write-Host "Building standalone executable..."
.\venv\Scripts\python.exe -m PyInstaller `
    --onefile `
    --noconsole `
    --name $AppName `
    --icon .\assets\app_icon.ico `
    --clean `
    .\pptx_to_text.py

Write-Host ""
Write-Host "Build complete:"
Write-Host "  .\dist\$AppName.exe"
