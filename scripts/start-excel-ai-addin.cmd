@echo off
setlocal
cd /d "%~dp0.." || exit /b 1
if not exist "logs" mkdir "logs"
if not exist "catalog" mkdir "catalog"
copy /y "manifest.xml" "catalog\manifest.xml" >nul

:restart
echo [%date% %time%] Starting Excel AI local servers >> "logs\startup.log"
call npm.cmd run all >> "logs\startup.log" 2>&1
echo [%date% %time%] Servers stopped; retrying in 5 seconds >> "logs\startup.log"
timeout /t 5 /nobreak >nul
goto restart
