<#
    Переключает рабочий сервер на выбранный выпуск.

    npm run release только выбирает выпуск для следующего запуска: работающий
    процесс намеренно удерживает свой, чтобы пересборка dist не подменяла
    проверенную панель под ногами. Из-за этого легко собрать выпуск и забыть
    перезапустить — панель продолжает отдавать старый код, а правки выглядят
    неработающими. Этот скрипт снимает процесс сервера; супервизор поднимает
    новый уже с выбранным выпуском.
#>
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$currentPath = Join-Path $root "releases\current.json"
if (-not (Test-Path $currentPath)) {
    Write-Error "Не найден releases\current.json: выпуск ещё не выбирался."
}
$selected = (Get-Content $currentPath -Raw | ConvertFrom-Json).id
Write-Host "Выбранный выпуск: $selected"

# Excel подключён через общий каталог catalog\, а он не хранится в git
# и раньше обновлялся руками. 19 сентября 2026 года это стоило лого:
# манифест в проекте получил новые иконки, а Excel читал старую копию
# из каталога и показывал заглушку. Копии сверяются при каждом выпуске.
$catalog = Join-Path $root "catalog"
if (Test-Path $catalog) {
    foreach ($name in @("manifest.xml", "manifest.dev.xml")) {
        $source = Join-Path $root $name
        $target = Join-Path $catalog $name
        if ((Test-Path $source) -and (-not (Test-Path $target) -or
            (Get-FileHash $source).Hash -ne (Get-FileHash $target).Hash)) {
            Copy-Item $source $target -Force
            Write-Host "Каталог обновлён: $name. Excel увидит изменения после перезапуска."
        }
    }
}

$server = Get-CimInstance Win32_Process -Filter "Name='node.exe'" |
    Where-Object { $_.CommandLine -like '*releases*server.js*' } |
    Select-Object -First 1

if (-not $server) {
    Write-Host "Сервер не запущен. Супервизор поднимет выбранный выпуск сам."
    exit 0
}

Write-Host "Останавливаю сервер, PID $($server.ProcessId)."
Stop-Process -Id $server.ProcessId -Force

# Проверяем по командной строке процесса, а не по HTTPS: Invoke-RestMethod
# в Windows PowerShell 5.1 не умеет -SkipCertificateCheck и спотыкается
# о самоподписанный сертификат localhost.
$deadline = (Get-Date).AddSeconds(60)
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Seconds 2
    $running = Get-CimInstance Win32_Process -Filter "Name='node.exe'" |
        Where-Object { $_.CommandLine -like '*releases*server.js*' } |
        Select-Object -First 1
    if ($running) {
        if ($running.CommandLine -like "*$selected*") {
            Write-Host "Сервер работает на выпуске $selected, PID $($running.ProcessId)."
            exit 0
        }
        Write-Host "Сервер поднялся, но на другом выпуске. Жду."
    }
}
Write-Error "Сервер не перешёл на выпуск $selected за отведённое время. Проверьте logs\server.log."
