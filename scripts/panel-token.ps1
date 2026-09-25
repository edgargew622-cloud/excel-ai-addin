<#
  Токен панели и манифест каталога с ним (этап 8, 8.0.1).

  Подключается точкой: . (Join-Path $PSScriptRoot 'panel-token.ps1')

  Токен — случайные 32 байта в server\panel-token (тот же формат пишет
  server/src/panelToken.ts: кто первым, тот и создаёт). В манифест каталога
  адрес панели попадает с ?t=<токен>; Excel открывает панель по нему, и
  панель прикладывает токен к запросам. Каталог и токен лежат в папке
  надстройки, закрытой установщиком для других пользователей компьютера.
#>

function Get-PanelToken([string] $Root) {
  $file = Join-Path $Root 'server\panel-token'
  if (-not (Test-Path -LiteralPath $file)) {
    $bytes = New-Object byte[] 32
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $rng.GetBytes($bytes)
    $rng.Dispose()
    $token = [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    [System.IO.File]::WriteAllText($file, $token, (New-Object System.Text.UTF8Encoding($false)))
  }
  $token = ([System.IO.File]::ReadAllText($file)).Trim()
  if ($token -notmatch '^[A-Za-z0-9_-]{32,}$') { throw "Файл токена панели повреждён: $file. Удалите его и запустите установку снова." }
  return $token
}

# Манифест для каталога: адрес панели с токеном. Остальные адреса (иконки,
# commands.html) токена не требуют: секретов в них нет.
function Get-CatalogManifest([string] $Root) {
  $token = Get-PanelToken $Root
  $xml = [System.IO.File]::ReadAllText((Join-Path $Root 'manifest.xml'))
  $withToken = $xml.Replace('/taskpane.html"', "/taskpane.html?t=$token`"")
  if ($withToken -eq $xml) { throw 'В manifest.xml не найден адрес панели taskpane.html.' }
  return $withToken
}

# Пишет catalog\manifest.xml, если он отличается. Возвращает $true, если записал.
function Write-CatalogManifest([string] $Root) {
  $catalog = Join-Path $Root 'catalog'
  New-Item -ItemType Directory -Path $catalog -Force | Out-Null
  $target = Join-Path $catalog 'manifest.xml'
  $content = Get-CatalogManifest $Root
  if ((Test-Path -LiteralPath $target) -and ([System.IO.File]::ReadAllText($target) -eq $content)) { return $false }
  [System.IO.File]::WriteAllText($target, $content, (New-Object System.Text.UTF8Encoding($false)))
  return $true
}
