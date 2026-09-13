$ErrorActionPreference = 'Stop'

$projectPath = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$catalogPath = Join-Path $projectPath 'catalog'
$manifestPath = Join-Path $projectPath 'manifest.xml'
$catalogManifestPath = Join-Path $catalogPath 'manifest.xml'

New-Item -ItemType Directory -Path $catalogPath -Force | Out-Null
Copy-Item -LiteralPath $manifestPath -Destination $catalogManifestPath -Force

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
