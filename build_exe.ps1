$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$distRoot = Join-Path $projectRoot "dist"
$buildRoot = Join-Path $projectRoot "build"
$releaseRoot = Join-Path $distRoot "Email Builder"
$zipPath = Join-Path $distRoot "Email Builder.zip"

Set-Location $projectRoot

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    throw "Python is required to build the exe."
}

python -m pip install PyInstaller openpyxl ttkbootstrap pywin32 | Out-Host

if (Test-Path $buildRoot) {
    Remove-Item -LiteralPath $buildRoot -Recurse -Force
}
if (Test-Path $releaseRoot) {
    Remove-Item -LiteralPath $releaseRoot -Recurse -Force
}
if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name "Email Builder" `
    --collect-all ttkbootstrap `
    --hidden-import win32timezone `
    --distpath $distRoot `
    email_template_gui.py

Copy-Item -LiteralPath (Join-Path $projectRoot "Reference Data") -Destination $releaseRoot -Recurse -Force
tar.exe -a -c -f $zipPath -C $distRoot "Email Builder"

Write-Host ""
Write-Host "Build complete:"
Write-Host "  $releaseRoot\Email Builder.exe"
Write-Host "  $zipPath"
