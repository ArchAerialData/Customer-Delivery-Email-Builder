@echo off
setlocal
cd /d "%~dp0"
if exist "%LocalAppData%\Programs\Python\Launcher\py.exe" (
    py -3 "%~dp0email_template_gui.py"
) else (
    python "%~dp0email_template_gui.py"
)

if errorlevel 1 (
    echo.
    echo Email Builder failed to start.
    echo.
    echo If the error mentions "Permission denied", close:
    echo "%~dp0Reference Data\XLSX Workbooks\Customer Information.xlsx"
    echo.
    pause
)

endlocal
