@echo off
rem Старые ярлыки используют этот путь; рабочим процессом управляет супервизор.
setlocal
cd /d "%~dp0.." || exit /b 1
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0start-server.ps1" %*
exit /b %ERRORLEVEL%
