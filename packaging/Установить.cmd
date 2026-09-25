@echo off
rem Double-click installer for the AI panel for Excel.
rem All messages are printed by PowerShell in Russian; this file stays ASCII
rem so that cmd.exe shows no garbled text on any code page.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\install.ps1" -Pause
exit /b %errorlevel%
