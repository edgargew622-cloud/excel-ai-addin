@echo off
rem Double-click uninstaller for the AI panel for Excel.
rem Messages are printed by PowerShell in Russian; this file stays ASCII.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\uninstall.ps1" -Pause
exit /b %errorlevel%
