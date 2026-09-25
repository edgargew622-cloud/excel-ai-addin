$ErrorActionPreference = 'Stop'

$projectPath = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$catalogPath = Join-Path $projectPath 'catalog'
# Только рабочий манифест: dev-надстройка на ленте не нужна (её убрали
# 19 сентября 2026 года), а в готовом комплекте manifest.dev.xml нет вовсе —
# прежде сценарий падал там на «Манифест не найден». В каталог он попадает
# с токеном панели в адресе (8.0.1).
$source = Join-Path $projectPath 'manifest.xml'
if (-not (Test-Path -LiteralPath $source)) { throw "Манифест не найден: $source" }
. (Join-Path $PSScriptRoot 'panel-token.ps1')
[void](Write-CatalogManifest $projectPath)

$driveRoot = [System.IO.Path]::GetPathRoot($catalogPath)
if ($driveRoot -notmatch '^[A-Za-z]:\\$') {
  throw "Unsupported catalog path: $catalogPath"
}
$driveLetter = $driveRoot.Substring(0, 1)
$relativePath = $catalogPath.Substring($driveRoot.Length)
$catalogUrl = '\\localhost\' + $driveLetter + '$\' + $relativePath

if (-not (Test-Path -LiteralPath (Join-Path $catalogUrl 'manifest.xml'))) {
  throw "Excel cannot read the local catalog: $catalogUrl"
}

$catalogId = '{7c7d2ddc-9f41-4d6c-a675-e073e604d789}'
$keyPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"
New-Item -Path $keyPath -Force | Out-Null
New-ItemProperty -Path $keyPath -Name Id -Value $catalogId -PropertyType String -Force | Out-Null
New-ItemProperty -Path $keyPath -Name Url -Value $catalogUrl -PropertyType String -Force | Out-Null
New-ItemProperty -Path $keyPath -Name Flags -Value 1 -PropertyType DWord -Force | Out-Null

Write-Output "Catalog registered: $catalogUrl"
