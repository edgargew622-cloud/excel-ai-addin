@echo off
rem Тонкая обёртка над супервизором. Оставлена ради существующих ярлыков:
rem прежний вариант этого файла перезапускал два процесса бесконечно, без
rem ограничения попыток и без проверки занятости порта. Вся логика запуска
rem теперь в scripts\start-server.ps1.
setlocal
cd /d "%~dp0.." || exit /b 1
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0start-server.ps1" %*
exit /b %ERRORLEVEL%
