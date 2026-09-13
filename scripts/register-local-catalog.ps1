$ErrorActionPreference = 'Stop'

$projectPath = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$catalogPath = Join-Path $projectPath 'catalog'
New-Item -ItemType Directory -Path $catalogPath -Force | Out-Null

# Рабочий и отладочный манифесты лежат в каталоге одновременно: у них разные
# Id, поэтому Excel показывает две кнопки и рабочая панель остаётся доступной,
# пока правится отладочная.
$manifests = @('manifest.xml', 'manifest.dev.xml')
foreach ($name in $manifests) {
  $source = Join-Path $projectPath $name
  if (-not (Test-Path -LiteralPath $source)) {
    Write-Warning "Манифест не найден, пропускаем: $name"
    continue
  }
  Copy-Item -LiteralPath $source -Destination (Join-Path $catalogPath $name) -Force
  Write-Output "Скопирован: $name"
}

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
