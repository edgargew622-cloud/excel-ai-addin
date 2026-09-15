[CmdletBinding()]
param([switch] $Rollback)

$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$releaseRoot = Join-Path $projectRoot 'releases'
$serverReleaseRoot = Join-Path $projectRoot 'server\releases'
$pointerPath = Join-Path $releaseRoot 'current.json'

function Assert-Release([string] $Id) {
  if ($Id -notmatch '^[0-9]{8}T[0-9]{6}Z-[a-f0-9]{8}$') { throw "Некорректный идентификатор выпуска: $Id" }
  $panel = Join-Path $releaseRoot "$Id\panel\taskpane.html"
  $server = Join-Path $serverReleaseRoot "$Id\dist\server.js"
  if (-not (Test-Path -LiteralPath $panel) -or -not (Test-Path -LiteralPath $server)) {
    throw "Выпуск $Id неполон. Текущий выпуск не изменён."
  }
}

function Set-Current([string] $Id, [string] $Previous) {
  Assert-Release $Id
  $temp = Join-Path $releaseRoot "current.$PID.tmp"
  $payload = @{ id = $Id; previous = $Previous; activatedAt = (Get-Date).ToUniversalTime().ToString('o') } |
    ConvertTo-Json -Compress
  [System.IO.File]::WriteAllText($temp, $payload, [System.Text.UTF8Encoding]::new($false))
  try {
    if (Test-Path -LiteralPath $pointerPath) {
      [System.IO.File]::Replace($temp, $pointerPath, (Join-Path $releaseRoot 'previous.json'))
    } else {
      [System.IO.File]::Move($temp, $pointerPath)
    }
  } finally {
    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force }
  }
}

$current = $null
if (Test-Path -LiteralPath $pointerPath) {
  $current = Get-Content -LiteralPath $pointerPath -Raw -Encoding UTF8 | ConvertFrom-Json
  Assert-Release ([string]$current.id)
}

if ($Rollback) {
  if (-not $current -or -not $current.previous) { throw 'Предыдущий выпуск отсутствует.' }
  Set-Current ([string]$current.previous) ([string]$current.id)
  Write-Output "Для следующего запуска выбран выпуск $($current.previous). Перезапустите свой сервер; открытые запросы завершите перед остановкой."
  exit 0
}

$panelSource = Join-Path $projectRoot 'dist'
$serverSource = Join-Path $projectRoot 'server\dist'
if (-not (Test-Path -LiteralPath (Join-Path $panelSource 'taskpane.html')) -or
    -not (Test-Path -LiteralPath (Join-Path $serverSource 'server.js'))) {
  throw 'Сборка не найдена. Сначала выполните npm run check.'
}

$stamp = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
$hash = (Get-FileHash -LiteralPath (Join-Path $serverSource 'server.js') -Algorithm SHA256).Hash.Substring(0,8).ToLowerInvariant()
$id = "$stamp-$hash"
$panelTarget = Join-Path $releaseRoot "$id\panel"
$serverTarget = Join-Path $serverReleaseRoot "$id\dist"
if ((Test-Path -LiteralPath $panelTarget) -or (Test-Path -LiteralPath $serverTarget)) { throw "Выпуск $id уже существует." }

New-Item -ItemType Directory -Path $panelTarget, $serverTarget -Force | Out-Null
Get-ChildItem -LiteralPath $panelSource -Force | Copy-Item -Destination $panelTarget -Recurse -Force
Get-ChildItem -LiteralPath $serverSource -Force | Copy-Item -Destination $serverTarget -Recurse -Force
Assert-Release $id
Set-Current $id $(if ($current) { [string]$current.id } else { '' })
Write-Output "Выпуск $id сохранён и выбран для следующего запуска."
if ($current) {
  Write-Output "Предыдущий выпуск $($current.id) сохранён. Для возврата: npm run rollback."
} else {
  Write-Output 'Это первый сохранённый выпуск; возврат станет доступен после следующего выпуска.'
}
