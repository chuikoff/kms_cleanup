@echo off
REM KMS Cleanup Tool - Windows 7
REM Right-click this file -> Run as administrator

cd /d "%~dp0"
echo [%date% %time%] Bat started > "%~dp0startup.log"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0kms_cleanup_win7.ps1" %* 1>> "%~dp0startup.log" 2>> "%~dp0error.log"
echo [%date% %time%] Exit code %ERRORLEVEL% >> "%~dp0startup.log"
if errorlevel 1 (
    echo.
    echo Error. Check these files in this folder:
    echo   error.log
    echo   kms_cleanup_win7_boot.log
    echo   kms_cleanup_win7.log
    pause
)