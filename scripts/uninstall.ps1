<#
  Удаление AI-панели для Excel. Обычно запускается двойным щелчком по
  «Удалить.cmd» в папке надстройки; вручную:

    powershell -ExecutionPolicy Bypass -File scripts\uninstall.ps1

  Что делает:
    1. останавливает сервер надстройки из этой папки;
    2. снимает автозапуск при входе в Windows;
    3. убирает регистрацию надстройки в Excel (каталог и запись «для разработчика»);
    4. удаляет из кэша Office записи этой надстройки;
    5. по вопросу — удаляет сертификат localhost и саму папку вместе с ключами.

  Трогает только то, что указывает на эту папку: чужие надстройки, задачи и
  каталоги остаются как были.
#>

[CmdletBinding()]
param(
  # Двойной щелчок: окно ждёт Enter и задаёт вопросы; без -Pause вопросов нет.
  [switch] $Pause,
  # Без вопросов удалить и папку (для проверок и сценариев).
  [switch] $RemoveFolder,
  [string] $TaskName = 'ExcelAiAddinServer'
)

$ErrorActionPreference = 'Stop'

function Ask([string] $Question) {
  if (-not $Pause) { return $false }
  $answer = Read-Host "$Question [д/Н]"
  return $answer -match '^\s*(д|да|y|yes)\s*$'
}

$exitCode = 0
try {
  $root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
  $manifest = Join-Path $root 'manifest.xml'
  if (-not (Test-Path -LiteralPath $manifest)) { throw "В $root нет manifest.xml — это не папка надстройки." }
  $addinId = [regex]::Match([System.IO.File]::ReadAllText($manifest), '<Id>([^<]+)</Id>').Groups[1].Value
  if (-not $addinId) { throw 'В manifest.xml не найден Id надстройки.' }
  $inRoot = { param([string] $path) $path -and ([System.IO.Path]::GetFullPath($path.Replace('/', '\')).TrimEnd('\') + '\').StartsWith($root.TrimEnd('\') + '\', [System.StringComparison]::OrdinalIgnoreCase) }

  Write-Output "Удаление AI-панели из $root"
  if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel открыт. Закройте его полностью и запустите удаление ещё раз: пока Excel работает, он держит надстройку.'
  }

  Write-Output '1/4 Остановка сервера'
  $processes = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object { ($_.Name -eq 'node.exe' -or $_.Name -eq 'powershell.exe') -and $_.CommandLine -and $_.ProcessId -ne $PID -and
      $_.CommandLine.IndexOf($root, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -and $_.CommandLine -notmatch 'uninstall\.ps1' })
  foreach ($process in $processes) { Stop-Process -Id $process.ProcessId -Force -ErrorAction SilentlyContinue }
  Write-Output "     остановлено процессов: $($processes.Count)"

  Write-Output '2/4 Автозапуск'
  $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
  if ($task -and @($task.Actions | Where-Object { & $inRoot ([regex]::Match("$($_.Arguments)", '-File\s+"?([^"]+\.ps1)').Groups[1].Value) }).Count) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Output "     задача $TaskName снята"
  } elseif ($task) {
    Write-Output "     задача $TaskName настроена на другую папку — не трогаем"
  } else {
    Write-Output '     автозапуска нет'
  }

  Write-Output '3/4 Регистрация в Excel'
  $wefKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF'
  $removed = 0
  $catalogs = Join-Path $wefKey 'TrustedCatalogs'
  if (Test-Path -LiteralPath $catalogs) {
    foreach ($item in Get-ChildItem -LiteralPath $catalogs) {
      $url = (Get-ItemProperty -LiteralPath $item.PSPath -ErrorAction SilentlyContinue).Url
      # Каталог записан как \\localhost\C$\путь — приводим к C:\путь.
      $local = if ($url -match '^\\\\localhost\\([A-Za-z])\$\\(.*)$') { "$($Matches[1]):\$($Matches[2])" } else { $url }
      if (& $inRoot $local) { Remove-Item -LiteralPath $item.PSPath -Recurse -Force; $removed++ }
    }
  }
  $developer = Join-Path $wefKey 'Developer'
  if (Test-Path -LiteralPath $developer) {
    $path = (Get-ItemProperty -LiteralPath $developer -ErrorAction SilentlyContinue).$addinId
    if (& $inRoot $path) {
      Remove-ItemProperty -LiteralPath $developer -Name $addinId -Force
      if (Test-Path -LiteralPath (Join-Path $developer $addinId)) { Remove-Item -LiteralPath (Join-Path $developer $addinId) -Recurse -Force }
      $removed++
    }
  }
  $wef = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
  $stale = @(Get-ChildItem -LiteralPath $wef -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "$addinId*" -and $_.FullName -notlike '*\webview2\*' })
  $stale | Remove-Item -Force -ErrorAction SilentlyContinue
  Write-Output "     регистраций снято: $removed; записей кэша Office удалено: $($stale.Count)"

  Write-Output '4/4 Сертификат и папка'
  $certCli = Join-Path $root 'server\node_modules\office-addin-dev-certs\lib\cli.js'
  $node = Join-Path $root 'node\node.exe'
  if (-not (Test-Path -LiteralPath $node)) { $node = 'node.exe' }
  if ((Test-Path -LiteralPath $certCli) -and (Ask 'Удалить сертификат localhost? Им пользуются и другие надстройки для разработчиков; если не знаете — ответьте «нет».')) {
    & $node $certCli uninstall
    Write-Output '     сертификат удалён'
  } else {
    Write-Output '     сертификат оставлен'
  }

  $removeFolder = $RemoveFolder -or (Ask "Удалить папку $root вместе с сохранёнными ключами провайдеров?")
  if ($removeFolder) {
    # Защита от ошибки: удаляем только папку, похожую на установку надстройки.
    $drive = [System.IO.Path]::GetPathRoot($root).TrimEnd('\')
    $isBundle = (Test-Path -LiteralPath (Join-Path $root 'bundle.json')) -or (Test-Path -LiteralPath (Join-Path $root 'releases\current.json'))
    if ($root.TrimEnd('\') -eq $drive -or $root.TrimEnd('\') -eq $env:USERPROFILE.TrimEnd('\') -or -not $isBundle) {
      throw "Папка $root не похожа на установку надстройки — удалите её вручную, если она не нужна."
    }
    # Из этой папки сейчас выполняется сам сценарий, а окно «Удалить.cmd»
    # держит её рабочей папкой: удаляем отдельным процессом, который ждёт,
    # пока закроются и сценарий, и окно.
    $parent = (Get-CimInstance Win32_Process -Filter "ProcessId=$PID" -ErrorAction SilentlyContinue).ParentProcessId
    $wait = (@($PID, $parent) | Where-Object { $_ }) -join ','
    $escaped = $root.Replace("'", "''")
    $command = "Wait-Process -Id $wait -ErrorAction SilentlyContinue; Start-Sleep -Seconds 1; Remove-Item -LiteralPath '$escaped' -Recurse -Force -ErrorAction SilentlyContinue"
    Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile', '-NonInteractive', '-WindowStyle', 'Hidden', '-Command', $command -WindowStyle Hidden -WorkingDirectory $env:TEMP
    Write-Output '     папка будет удалена сразу после закрытия этого окна'
  } else {
    Write-Output "     папка оставлена: $root — её можно удалить вручную"
  }

  Write-Output ''
  Write-Output 'Готово: надстройка удалена. При следующем запуске Excel её на ленте не будет.'
} catch {
  $exitCode = 1
  Write-Host ''
  Write-Host "Удаление не завершено: $($_.Exception.Message)" -ForegroundColor Red
  if (-not $Pause) { throw }
} finally {
  if ($Pause) {
    Write-Host ''
    [void](Read-Host 'Нажмите Enter, чтобы закрыть окно')
  }
}
exit $exitCode
