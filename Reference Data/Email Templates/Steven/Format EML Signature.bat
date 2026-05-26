@echo off
setlocal

if "%~1"=="" (
    echo Drag and drop one or more .eml files onto this file to format them in place.
    pause
    exit /b 1
)

python "%~dp0format_eml_signature.py" %*
if errorlevel 1 (
    echo.
    pause
    exit /b 1
)

echo.
echo Done.
pause
