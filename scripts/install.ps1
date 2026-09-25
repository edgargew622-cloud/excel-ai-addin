<#
  Установка AI-панели для Excel из готового комплекта — одной командой:

    powershell -ExecutionPolicy Bypass -File scripts\install.ps1

  Шаги:
    1. сертификат для https://localhost — Windows спросит, доверять ли ему;
    2. автозапуск сервера при входе в Windows;
    3. запуск сервера сейчас и проверка, что отвечает именно этот выпуск;
    4. регистрация надстройки в Excel (общая папка надёжных надстроек).

  Повторный запуск безопасен: каждый шаг либо уже сделан, либо делается заново.
  Ключи провайдеров вводятся потом в самой панели кнопкой «Ключи».
#>

[CmdletBinding()]
param(
  # В CI сертификат ставится для всей машины заранее, без окна подтверждения.
  [switch] $SkipCertificate,
  [string] $TaskName = 'ExcelAiAddinServer'
)

# Манифест надстройки указывает на https://localhost:3000 — другой порт Excel не найдёт.
$Port = 3000

$ErrorActionPreference = 'Stop'

$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$bundledNode = Join-Path $root 'node\node.exe'
$node = if (Test-Path -LiteralPath $bundledNode) { $bundledNode } else { 'node.exe' }
$pointer = Join-Path $root 'releases\current.json'
if (-not (Test-Path -LiteralPath $pointer)) {
  throw "Не найден $pointer. Запускайте установку из распакованного комплекта, а в рабочей копии сначала выполните npm run release."
}
$release = (Get-Content -LiteralPath $pointer -Raw | ConvertFrom-Json).id

function Get-Health {
  try {
    return Invoke-RestMethod -Uri "https://localhost:$Port/api/health" -TimeoutSec 3
  } catch {
    return $null
  }
}

function Wait-Health([int] $Seconds) {
  $deadline = (Get-Date).AddSeconds($Seconds)
  while ((Get-Date) -lt $deadline) {
    $health = Get-Health
    if ($health) { return $health }
    Start-Sleep -Seconds 1
  }
  return $null
}

Write-Output "Установка AI-панели из $root (выпуск $release)"

Write-Output '1/4 Сертификат для https://localhost'
if ($SkipCertificate) {
  Write-Output '     пропущен по параметру -SkipCertificate'
} else {
  $cli = Join-Path $root 'server\node_modules\office-addin-dev-certs\lib\cli.js'
  if (-not (Test-Path -LiteralPath $cli)) { throw "Не найден установщик сертификата: $cli" }
  Write-Output '     Если Windows спросит, установить ли сертификат, ответьте «Да»: без него Excel покажет пустую панель.'
  & $node $cli install
  if ($LASTEXITCODE -ne 0) { throw 'Сертификат не установлен. Запустите установку ещё раз и подтвердите установку сертификата.' }
}

Write-Output '2/4 Автозапуск сервера при входе в Windows'
& (Join-Path $PSScriptRoot 'register-autostart.ps1') -TaskName $TaskName

Write-Output '3/4 Запуск сервера'
$health = Get-Health
if (-not $health) {
  Start-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
  $health = Wait-Health 20
}
if (-not $health) {
  # Задача планировщика может не стартовать вне интерактивного входа (так в
  # CI). Супервизор держит один экземпляр через именованный мьютекс, поэтому
  # прямой запуск не создаст второго сервера.
  Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList @(
    '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', "`"$(Join-Path $PSScriptRoot 'start-server.ps1')`""
  ) | Out-Null
  $health = Wait-Health 60
}
if (-not $health) {
  throw "Сервер не ответил на https://localhost:$Port. Подробности — в $(Join-Path $root 'logs\server.log') и $(Join-Path $root 'logs\app.log')."
}
if ($health.app -ne 'excel-ai-addin') {
  throw "Порт $Port занят другой программой. Освободите его и запустите установку снова."
}
if ($health.release -ne $release) {
  throw "На порту $Port уже работает другая копия надстройки (выпуск $($health.release)). Остановите её и запустите установку снова."
}
Write-Output "     сервер отвечает, выпуск $($health.release)"

Write-Output '4/4 Регистрация надстройки в Excel'
& (Join-Path $PSScriptRoot 'register-local-catalog.ps1')

Write-Output ''
Write-Output 'Готово. Дальше в Excel:'
Write-Output '  1. Полностью закройте и снова откройте Excel.'
Write-Output '  2. Главная → Надстройки → Мои надстройки → Общая папка → AI-панель → Добавить.'
Write-Output '  3. Нажмите «Открыть чат», затем «Ключи» — и вставьте ключ провайдера.'
