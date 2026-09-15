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
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
  $ownsTask = @($existing.Actions | Where-Object { $_.Execute -like '*powershell*' -and $_.Arguments -like "*$supervisor*" }).Count -gt 0
  if (-not $ownsTask) { throw "Задача $TaskName уже принадлежит другой программе; не изменяем её." }
}

if ($Remove) {
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

# Сохранить старые ярлыки вместо удаления. Трогаем только ссылки, ведущие на
# локальный legacy launcher именно этого проекта.
$startupDir = [Environment]::GetFolderPath('Startup')
$startupFull = [System.IO.Path]::GetFullPath($startupDir).TrimEnd('\')
$legacyLauncher = [System.IO.Path]::GetFullPath((Join-Path $projectRoot 'scripts\start-excel-ai-addin.cmd'))
$backupDir = Join-Path $projectRoot 'releases\disabled-startup-links'
$shell = New-Object -ComObject WScript.Shell
foreach ($item in (Get-ChildItem -LiteralPath $startupDir -Filter '*.lnk' -File -ErrorAction SilentlyContinue)) {
  $sourceFull = [System.IO.Path]::GetFullPath($item.FullName)
  if (-not $sourceFull.StartsWith($startupFull + '\', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
  $shortcut = $shell.CreateShortcut($sourceFull)
  if ([string]::IsNullOrWhiteSpace($shortcut.TargetPath)) { continue }
  if (-not [string]::Equals([System.IO.Path]::GetFullPath($shortcut.TargetPath), $legacyLauncher, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
  New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
  $destination = Join-Path $backupDir $item.Name
  if (Test-Path -LiteralPath $destination) { throw "Резервная копия ярлыка уже существует: $destination" }
  Move-Item -LiteralPath $sourceFull -Destination $destination
  Write-Output "Прежний ярлык сохранён: $destination"
}

Write-Output "Автозапуск зарегистрирован: задача $TaskName при входе пользователя $env:USERNAME."
Write-Output "Проверить сейчас:  Start-ScheduledTask -TaskName $TaskName"
Write-Output "Состояние:         Get-ScheduledTask -TaskName $TaskName | Get-ScheduledTaskInfo"
Write-Output "Диагностика:       scripts\diagnose.ps1"
