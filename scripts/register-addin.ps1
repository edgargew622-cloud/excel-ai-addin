<#
  Регистрация надстройки в Excel (этап 8, 8.8.2).

  Основной способ — каталог надёжных надстроек («Общая папка» в Excel) по
  адресу \\localhost\<диск>$\…\catalog. Служебная общая папка диска открыта
  только администраторам, поэтому у пользователя без этих прав каталог
  недоступен.

  Запасной способ — запись в HKCU\…\WEF\Developer, куда писать может любой
  пользователь. Проверка 27.09.2026 показала: Excel такую надстройку сам не
  показывает — ни на ленте, ни в «Моих надстройках»; он подхватывает её только
  из книги, в которую она вставлена. Поэтому запись остаётся лишь запасным
  вариантом, и установка честно предупреждает о нём.

  Манифест берётся из catalog\manifest.xml: в нём адрес панели с токеном
  (8.0.1), и лежит он в закрытой папке надстройки. Регистрация другим
  способом на эту же папку снимается, чтобы надстройка не появилась дважды.
#>

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
$catalogReadable = $catalogUrl -and (Test-Path -LiteralPath (Join-Path $catalogUrl 'manifest.xml') -ErrorAction SilentlyContinue)

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

if ($catalogReadable) {
  $keyPath = Join-Path $catalogs $catalogId
  New-Item -Path $keyPath -Force | Out-Null
  New-ItemProperty -Path $keyPath -Name Id -Value $catalogId -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $keyPath -Name Url -Value $catalogUrl -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $keyPath -Name Flags -Value 1 -PropertyType DWord -Force | Out-Null
  if ((Get-ItemProperty -LiteralPath $developer -ErrorAction SilentlyContinue).$addinId) {
    Remove-ItemProperty -LiteralPath $developer -Name $addinId -Force
  }
  Write-Output "Надстройка зарегистрирована в общей папке: $catalogUrl"
} else {
  if (Test-Path -LiteralPath (Join-Path $catalogs $catalogId)) { Remove-Item -LiteralPath (Join-Path $catalogs $catalogId) -Recurse -Force }
  New-Item -Path $developer -Force | Out-Null
  New-ItemProperty -Path $developer -Name $addinId -Value $manifest -PropertyType String -Force | Out-Null
  Write-Warning "Общая папка $catalogUrl недоступна (у этой учётной записи Windows нет прав администратора). Надстройка записана как «для разработчика»: Excel может её не показать."
  Write-Output "Надстройка зарегистрирована для разработчика: $manifest"
}
