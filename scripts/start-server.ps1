<#
  Супервизор рабочего сервера надстройки.

  Заменяет бесконечный goto restart из прежнего .cmd: тот перезапускал связку
  без ограничения попыток, без нарастающей паузы и без проверки занятости порта,
  поэтому при устойчивой ошибке писал в журнал до конца диска.

  Правила:
    занятый порт проверяется по идентификатору в /api/health — чужой процесс
      не завершается никогда, о нём сообщается;
    попытки ограничены, пауза между ними растёт;
    процесс, проживший достаточно долго, считается новым случаем, и счётчик
      попыток сбрасывается: это авария, а не цикл перезапуска;
    журналы ротируются по размеру.
#>

[CmdletBinding()]
param(
  [int] $Port = 3000,
  [int] $MaxAttempts = 5,
  [int] $StableRunSeconds = 60,
  [int] $MaxLogBytes = 5MB,
  [int] $KeepLogs = 3
)

$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$logDir = Join-Path $projectRoot 'logs'
$logPath = Join-Path $logDir 'server.log'
$entryPoint = Join-Path $projectRoot 'server\dist\server.js'
$pointerPath = Join-Path $projectRoot 'releases\current.json'
$appId = 'excel-ai-addin'

# Готовый комплект несёт свой Node в node\node.exe: пользователю не нужно
# ставить Node.js. В рабочей копии разработчика используется Node из PATH.
$bundledNode = Join-Path $projectRoot 'node\node.exe'
$nodeExe = if (Test-Path -LiteralPath $bundledNode) { $bundledNode } else { 'node.exe' }

if (-not (Test-Path -LiteralPath $logDir)) {
  New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

function Write-Log {
  param([string] $Message)
  $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
  Add-Content -LiteralPath $logPath -Value $line -Encoding utf8
  Write-Output $line
}

function Rotate-Log {
  if (-not (Test-Path -LiteralPath $logPath)) { return }
  $size = (Get-Item -LiteralPath $logPath).Length
  if ($size -lt $MaxLogBytes) { return }

  for ($i = $KeepLogs - 1; $i -ge 1; $i--) {
    $from = "$logPath.$i"
    $to = "$logPath.$($i + 1)"
    if (Test-Path -LiteralPath $from) {
      Move-Item -LiteralPath $from -Destination $to -Force
    }
  }
  Move-Item -LiteralPath $logPath -Destination "$logPath.1" -Force
}

<#
  Кто занимает порт. Возвращает 'free', 'ours' или 'foreign'.
  Собственный экземпляр узнаём по полю app в /api/health, а не по факту,
  что порт отвечает: отвечать может любая программа.
#>
function Get-PortOwner {
  $probe = $null
  try {
    $probe = Invoke-RestMethod -Uri "https://127.0.0.1:$Port/api/health" -TimeoutSec 5
  } catch {
    # Порт может быть занят программой, которая не отвечает на этот путь.
    $busy = $false
    try {
      $busy = [bool](Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction Stop)
    } catch {
      $busy = $false
    }
    if ($busy) { return 'foreign' }
    return 'free'
  }

  if ($null -ne $probe -and $probe.app -eq $appId) { return 'ours' }
  return 'foreign'
}

# Проверки занятости порта недостаточно: между падением сервера и его подъёмом
# есть окно, в которое второй супервизор проскочит и начнёт свой цикл
# перезапусков параллельно первому. Именованный мьютекс закрывает это окно.
$mutex = New-Object System.Threading.Mutex($false, 'Global\ExcelAiAddinSupervisor')
$holdsMutex = $false
try {
  $holdsMutex = $mutex.WaitOne(0)
} catch [System.Threading.AbandonedMutexException] {
  # Предыдущий супервизор завершился, не освободив мьютекс: владение переходит
  # к нам, и это нормальная ситуация после аварийного снятия процесса.
  $holdsMutex = $true
}
if (-not $holdsMutex) {
  Write-Log 'Супервизор уже запущен в этой системе. Второй экземпляр не нужен.'
  exit 0
}

try {

Rotate-Log

function Select-Release {
  if (-not (Test-Path -LiteralPath $pointerPath)) {
    throw 'Рабочий выпуск не выбран. Выполните npm run check, затем npm run release.'
  }
  $selected = Get-Content -LiteralPath $pointerPath -Raw -Encoding UTF8 | ConvertFrom-Json
  $id = [string]$selected.id
  if ($id -notmatch '^[0-9]{8}T[0-9]{6}Z-[a-f0-9]{8}$') { throw "Некорректный идентификатор выпуска: $id" }
  $script:entryPoint = Join-Path $projectRoot "server\releases\$id\dist\server.js"
  $panelPath = Join-Path $projectRoot "releases\$id\panel"
  if (-not (Test-Path -LiteralPath $entryPoint) -or -not (Test-Path -LiteralPath (Join-Path $panelPath 'taskpane.html'))) {
    throw "Выпуск $id неполон. Сервер не запущен."
  }
  $env:EXCEL_AI_PROJECT_ROOT = $projectRoot
  $env:PANEL_DIST_DIR = $panelPath
  $env:EXCEL_AI_RELEASE_ID = $id
  return $id
}

$owner = Get-PortOwner
if ($owner -eq 'ours') {
  Write-Log "Порт $Port уже держит рабочий экземпляр надстройки. Второй не запускаем."
  exit 0
}
if ($owner -eq 'foreign') {
  Write-Log "Порт $Port занят другой программой. Она не будет завершена — освободите порт вручную или измените PORT в server/.env."
  exit 3
}

$attempt = 0
while ($attempt -lt $MaxAttempts) {
  $attempt++
  Rotate-Log
  $selectedId = Select-Release
  Write-Log "Запуск выпуска $selectedId, попытка $attempt из $MaxAttempts."

  $startedAt = Get-Date
  # -Wait вместе с -PassThru: только так ExitCode заполняется надёжно. С
  # отдельным WaitForExit код выхода возвращался пустым, и проверка кода 10
  # («порт занят, повторять бессмысленно») молча никогда не срабатывала.
  # Приложение само пишет ограниченный logs/app.log; перенаправление stdout
  # в файл оставило бы его без ротации до завершения процесса.
  $process = Start-Process -FilePath $nodeExe -ArgumentList $entryPoint `
    -WorkingDirectory $projectRoot -WindowStyle Hidden -PassThru -Wait

  $ranSeconds = ((Get-Date) - $startedAt).TotalSeconds
  $code = $process.ExitCode

  # Код -1 означает завершение извне (TerminateProcess), а не аварию
  # приложения: так выглядит снятие процесса пользователем или инструментом.
  if ($null -eq $code) {
    $codeText = 'не получен'
  } elseif ($code -eq -1) {
    $codeText = '-1, процесс снят извне'
  } else {
    $codeText = $code
  }
  Write-Log ("Сервер остановлен: код {0}, проработал {1:N0} с." -f $codeText, $ranSeconds)

  # Код 10 означает занятый порт: сервер сам отказался поднимать второй
  # экземпляр, и повторять это бессмысленно.
  if ($code -eq 10) {
    Write-Log 'Порт занят по сообщению самого сервера. Перезапуск не имеет смысла.'
    exit 10
  }

  if ($ranSeconds -ge $StableRunSeconds) {
    Write-Log 'Процесс проработал достаточно долго: считаем это отдельной аварией, счётчик попыток сброшен.'
    $attempt = 0
    $delay = 5
  } else {
    $delay = [Math]::Min(5 * [Math]::Pow(2, $attempt - 1), 120)
  }

  if ($attempt -ge $MaxAttempts) { break }
  Write-Log ("Повтор через {0:N0} с." -f $delay)
  Start-Sleep -Seconds $delay
}

Write-Log "Исчерпаны $MaxAttempts попыток подряд. Супервизор остановлен, чтобы не писать в журнал бесконечно. Причину смотрите выше в $logPath."
exit 1

} finally {
  if ($holdsMutex) { $mutex.ReleaseMutex() }
  $mutex.Dispose()
}
