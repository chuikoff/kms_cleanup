@echo off
REM KMS Cleanup Tool - Windows 7
REM Right-click this file -> Run as administrator

cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0kms_cleanup_win7.ps1" %*
if errorlevel 1 pause