<#
  Регистрация надстройки в Excel без прав администратора (этап 8, 8.8.2).

  Прежде надстройка подключалась через каталог надёжных надстроек по адресу
  \\localhost\C$\… — служебная общая папка диска открыта только
  администраторам, и у пользователя без этих прав установка останавливалась.
  Теперь манифест регистрируется в HKCU\…\WEF\Developer — разделе текущего
  пользователя, куда запись разрешена всем. Так же Excel подключал надстройку
  при разработке.

  Манифест берётся из catalog\manifest.xml: в нём адрес панели с токеном
  (8.0.1), и лежит он в закрытой папке надстройки. Прежняя регистрация через
  каталог, указывающая на эту же папку, снимается, чтобы надстройка не
  появилась в Excel дважды.
#>

$ErrorActionPreference = 'Stop'

$projectPath = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
. (Join-Path $PSScriptRoot 'panel-token.ps1')
[void](Write-CatalogManifest $projectPath)
$manifest = Join-Path $projectPath 'catalog\manifest.xml'
$addinId = [regex]::Match([System.IO.File]::ReadAllText($manifest), '<Id>([^<]+)</Id>').Groups[1].Value
if (-not $addinId) { throw 'В manifest.xml не найден Id надстройки.' }

$developer = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
New-Item -Path $developer -Force | Out-Null
New-ItemProperty -Path $developer -Name $addinId -Value $manifest -PropertyType String -Force | Out-Null

# Прежний способ — каталог \\localhost\C$\<эта папка>\catalog. Снимаем только его.
$catalogs = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs'
if (Test-Path -LiteralPath $catalogs) {
  $ownCatalog = Join-Path $projectPath 'catalog'
  foreach ($item in Get-ChildItem -LiteralPath $catalogs) {
    $url = (Get-ItemProperty -LiteralPath $item.PSPath -ErrorAction SilentlyContinue).Url
    $local = if ($url -match '^\\\\localhost\\([A-Za-z])\$\\(.*)$') { "$($Matches[1]):\$($Matches[2])" } else { $url }
    if ($local -and $local.TrimEnd('\') -ieq $ownCatalog.TrimEnd('\')) {
      Remove-Item -LiteralPath $item.PSPath -Recurse -Force
      Write-Output "Прежняя регистрация через каталог снята: $url"
    }
  }
}

Write-Output "Надстройка зарегистрирована: $manifest"
