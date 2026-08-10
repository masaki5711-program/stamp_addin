@echo off
rem Date-stamp add-in uninstaller (double-click to run)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
pause
