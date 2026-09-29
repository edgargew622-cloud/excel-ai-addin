<#
  Регистрация надстройки в Excel (этап 8, 8.8.2).

  Основной способ — каталог надёжных надстроек («Общая папка» в Excel) по
  адресу \\localhost\<диск>$\…\catalog. Служебная общая папка диска открыта
  только администраторам, поэтому у пользователя без этих прав каталог
  недоступен.

  Без прав администратора (29.09.2026) — тот же каталог через WebDAV:
  сервер надстройки отдаёт папку catalog на http://127.0.0.1:3080 только для
  чтения, и Windows открывает её как сетевую папку \\localhost@3080\catalog
  (служба WebClient). Проверено: Excel нашёл надстройку и открыл панель.

  Крайний случай — запись в HKCU\…\WEF\Developer, куда писать может любой
  пользователь. Проверка 27.09.2026 показала: Excel такую надстройку сам не
  показывает — ни на ленте, ни в «Моих надстройках»; он подхватывает её только
  из книги, в которую она вставлена. Поэтому запись остаётся лишь запасным
  вариантом, и установка честно предупреждает о нём.

  Манифест берётся из catalog\manifest.xml: в нём адрес панели с токеном
  (8.0.1), и лежит он в закрытой папке надстройки. Регистрация другим
  способом на эту же папку снимается, чтобы надстройка не появилась дважды.
#>

# -SkipAdminShare: не пробовать \\localhost\C$ — так путь без прав администратора
# проверяется и на компьютере, где они есть (проверка комплекта в CI, ручная проверка).
param([switch] $SkipAdminShare)

$ErrorActionPreference = 'Stop'

$projectPath = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
. (Join-Path $PSScriptRoot 'panel-token.ps1')
[void](Write-CatalogManifest $projectPath)
$catalogPath = Join-Path $projectPath 'catalog'
$manifest = Join-Path $catalogPath 'manifest.xml'
$addinId = [regex]::Match([System.IO.File]::ReadAllText($manifest), '<Id>([^<]+)</Id>').Groups[1].Value
if (-not $addinId) { throw 'В manifest.xml не найден Id надстройки.' }

$wefKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF'
$catalogs = Join-Path $wefKey 'TrustedCatalogs'
$developer = Join-Path $wefKey 'Developer'
# Постоянный Id каталога: повторная установка перезаписывает ту же запись.
$catalogId = '{7c7d2ddc-9f41-4d6c-a675-e073e604d789}'

$driveRoot = [System.IO.Path]::GetPathRoot($catalogPath)
$catalogUrl = if ($driveRoot -match '^([A-Za-z]):\\$') { '\\localhost\' + $Matches[1] + '$\' + $catalogPath.Substring($driveRoot.Length) } else { $null }
$catalogReadable = -not $SkipAdminShare -and $catalogUrl -and (Test-Path -LiteralPath (Join-Path $catalogUrl 'manifest.xml') -ErrorAction SilentlyContinue)

# WebDAV-каталог сервера (catalogDav.ts). Порт — CATALOG_DAV_PORT из server\.env,
# иначе 3080. Сервер к этому моменту запущен (install.ps1, шаг 4); первое
# обращение будит службу WebClient, поэтому несколько попыток.
$davPort = 3080
$envFile = Join-Path $projectPath 'server\.env'
if (Test-Path -LiteralPath $envFile) {
  $line = Get-Content -LiteralPath $envFile | Where-Object { $_ -match '^\s*CATALOG_DAV_PORT\s*=' } | Select-Object -Last 1
  if ($line -match '=\s*(\d+)\s*$') { $davPort = [int]$Matches[1] }
}
$davUrl = "\\localhost@$davPort\catalog"
# Метка разрешает серверу отдавать каталог по WebDAV. Без неё манифест с
# токеном по WebDAV не виден (при регистрации через C$ он и не нужен).
# Папка server закрыта правами — другая учётная запись метку не поставит.
$davMarker = Join-Path $projectPath 'server\catalog-dav'
$davReadable = $false
if (-not $catalogReadable) {
  Set-Content -LiteralPath $davMarker -Value 'Каталог надстроек по WebDAV: регистрация без прав администратора (register-addin.ps1).' -Encoding UTF8
  for ($try = 1; $try -le 5 -and -not $davReadable; $try++) {
    $davReadable = Test-Path -LiteralPath (Join-Path $davUrl 'manifest.xml') -ErrorAction SilentlyContinue
    if (-not $davReadable) { Start-Sleep -Seconds 2 }
  }
}
if (-not $davReadable -and (Test-Path -LiteralPath $davMarker)) { Remove-Item -LiteralPath $davMarker -Force }

# Прежние каталоги, указывающие на эту же папку (под любым Id), снимаем.
if (Test-Path -LiteralPath $catalogs) {
  foreach ($item in Get-ChildItem -LiteralPath $catalogs) {
    $url = (Get-ItemProperty -LiteralPath $item.PSPath -ErrorAction SilentlyContinue).Url
    $local = if ($url -match '^\\\\localhost\\([A-Za-z])\$\\(.*)$') { "$($Matches[1]):\$($Matches[2])" } else { $url }
    if ($local -and $local.TrimEnd('\') -ieq $catalogPath.TrimEnd('\') -and $item.PSChildName -ne $catalogId) {
      Remove-Item -LiteralPath $item.PSPath -Recurse -Force
    }
  }
}

if ($catalogReadable -or $davReadable) {
  $url = if ($catalogReadable) { $catalogUrl } else { $davUrl }
  $keyPath = Join-Path $catalogs $catalogId
  New-Item -Path $keyPath -Force | Out-Null
  New-ItemProperty -Path $keyPath -Name Id -Value $catalogId -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $keyPath -Name Url -Value $url -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $keyPath -Name Flags -Value 1 -PropertyType DWord -Force | Out-Null
  if ((Get-ItemProperty -LiteralPath $developer -ErrorAction SilentlyContinue).$addinId) {
    Remove-ItemProperty -LiteralPath $developer -Name $addinId -Force
  }
  if ($catalogReadable) { Write-Output "Надстройка зарегистрирована в общей папке: $catalogUrl" }
  else { Write-Output "Надстройка зарегистрирована в общей папке через WebDAV (без прав администратора): $davUrl" }
} else {
  if (Test-Path -LiteralPath (Join-Path $catalogs $catalogId)) { Remove-Item -LiteralPath (Join-Path $catalogs $catalogId) -Recurse -Force }
  New-Item -Path $developer -Force | Out-Null
  New-ItemProperty -Path $developer -Name $addinId -Value $manifest -PropertyType String -Force | Out-Null
  Write-Warning "Общая папка недоступна: $catalogUrl требует прав администратора, а $davUrl не открылся (служба Windows WebClient выключена или сервер не запущен). Надстройка записана как «для разработчика»: Excel может её не показать."
  Write-Output "Надстройка зарегистрирована для разработчика: $manifest"
}
