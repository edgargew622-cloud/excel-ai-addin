<#
  Автозапуск рабочего сервера при входе пользователя в Windows.

  Задача планировщика вместо ярлыка в автозагрузке: она не зависит от окна
  терминала, переживает выход из системы и её состояние видно в планировщике.

  Сценарий заодно убирает прежний ярлык автозагрузки и прежний .cmd, если они
  были: иначе два механизма создадут два экземпляра сервера.

  Снять автозапуск: scripts\register-autostart.ps1 -Remove
#>

[CmdletBinding()]
param(
  [switch] $Remove,
  [string] $TaskName = 'ExcelAiAddinServer'
)

$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$supervisor = Join-Path $projectRoot 'scripts\start-server.ps1'

if ($Remove) {
  $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
  if ($existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Output "Автозапуск снят: задача $TaskName удалена."
  } else {
    Write-Output "Задача $TaskName не зарегистрирована — снимать нечего."
  }
  exit 0
}

if (-not (Test-Path -LiteralPath $supervisor)) {
  throw "Не найден супервизор: $supervisor"
}

# Прежние механизмы запуска убираем, чтобы не получить два сервера сразу.
$startupDir = [Environment]::GetFolderPath('Startup')
foreach ($stale in @('excel-ai-addin.lnk', 'start-excel-ai-addin.lnk')) {
  $path = Join-Path $startupDir $stale
  if (Test-Path -LiteralPath $path) {
    Remove-Item -LiteralPath $path -Force
    Write-Output "Удалён прежний ярлык автозагрузки: $stale"
  }
}

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
  -Argument "-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$supervisor`"" `
  -WorkingDirectory $projectRoot

$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME

# RestartCount не задаём: перезапуском управляет сам супервизор, у него
# ограниченное число попыток и нарастающая пауза. Два механизма перезапуска
# поверх друг друга дали бы бесконечный цикл, от которого мы и уходим.
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
  -StartWhenAvailable -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
  -Description 'Локальный сервер AI-панели для Excel: раздаёт собранную панель и /api на порту 3000.' `
  -Force | Out-Null

Write-Output "Автозапуск зарегистрирован: задача $TaskName при входе пользователя $env:USERNAME."
Write-Output "Проверить сейчас:  Start-ScheduledTask -TaskName $TaskName"
Write-Output "Состояние:         Get-ScheduledTask -TaskName $TaskName | Get-ScheduledTaskInfo"
Write-Output "Диагностика:       scripts\diagnose.ps1"
