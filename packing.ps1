# Set script to exit on error
$ErrorActionPreference = "Stop"

# Script Title
Write-Host "MOD43 Calculator Build Script" -ForegroundColor Green

# Create virtual environment if it doesn't exist
$venvPath = "venv"
if (-not (Test-Path "$venvPath\Scripts\Activate.ps1")) {
    Write-Host "Creating virtual environment..." -ForegroundColor Yellow
    python -m venv $venvPath
    Write-Host "Virtual environment created." -ForegroundColor Green
}

# Clean old build files
Write-Host "Cleaning old build files..." -ForegroundColor Yellow
if (Test-Path "build") {
    Remove-Item -Path "build" -Recurse -Force -ErrorAction SilentlyContinue
}
if (Test-Path "dist") {
    # Remove entire dist directory to ensure clean build
    Remove-Item -Path "dist" -Recurse -Force -ErrorAction SilentlyContinue
}
if (Test-Path "MOD43_GUI.spec") {
    Remove-Item -Path "MOD43_GUI.spec" -Force -ErrorAction SilentlyContinue
}
# Clean PyInstaller cache
$cachePath = "$env:LOCALAPPDATA\pyinstaller"
if (Test-Path $cachePath) {
    Write-Host "Cleaning PyInstaller cache..." -ForegroundColor Yellow
    Remove-Item -Path $cachePath -Recurse -Force -ErrorAction SilentlyContinue
}

# Install dependencies
Write-Host "Installing dependencies..." -ForegroundColor Yellow
& "$venvPath\Scripts\pip.exe" install pyinstaller openpyxl | Out-Null

# Start packaging
Write-Host "Starting packaging..." -ForegroundColor Yellow
Write-Host "Creating executable..." -ForegroundColor Yellow

# Execute PyInstaller
$upxPath = "D:\Program Files\Python311\Scripts"
$pyinstaller_path = "$PWD\$venvPath\Scripts\pyinstaller.exe"
Write-Host "Executing PyInstaller..." -ForegroundColor Yellow

# Use direct command line
$command = "`"$pyinstaller_path`" --noconfirm --onefile --windowed --name=MOD43_GUI --icon=app_ico.ico --version-file=version_info.txt --add-data=app_ico.ico;. --upx-dir=`"$upxPath`" --upx-exclude=vcruntime140.dll --optimize=2 MOD43_GUI.py"

# Execute command
cmd /c $command

# Check if build succeeded
if (Test-Path "dist\MOD43_GUI.exe") {
    Write-Host "Packaging completed!" -ForegroundColor Green
    Write-Host "Executable file located at dist\MOD43_GUI.exe" -ForegroundColor Green
    Write-Host "Build successful!" -ForegroundColor Green
    Write-Host "Executable: $PWD\dist\MOD43_GUI.exe" -ForegroundColor Green
} else {
    Write-Host "Packaging failed!" -ForegroundColor Red
    exit 1
}

# Exit script
exit 0