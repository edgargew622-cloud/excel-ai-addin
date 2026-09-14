<#
  Подключение надстройки к Excel через developer-регистрацию в реестре.

  Это основной способ на этой машине. Каталог надёжных надстроек по UNC-пути
  (register-local-catalog.ps1) остаётся как запасной, но он показывает надстройку
  только в «Мои надстройки → Общая папка», тогда как developer-регистрация даёт
  кнопку на ленте сразу при запуске Excel.

  Проверить состояние:  scripts\register-addin.ps1 -Show
  Снять регистрацию:    scripts\register-addin.ps1 -Remove
#>

[CmdletBinding()]
param(
  [switch] $Remove,
  [switch] $Show,
  # Кэш ленты и сведений о надстройках. Excel держит его между запусками, и
  # вернувшаяся кнопка может не появиться, пока кэш не сброшен.
  [switch] $ClearCache,
  # По умолчанию подключаем только рабочий манифест. Отладочный добавляется
  # явно: он ссылается на порт 3100, который при обычной работе не запущен.
  [switch] $IncludeDev
)

$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$developerKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'

function Get-ManifestId {
  param([string] $Path)
  $xml = New-Object System.Xml.XmlDocument
  $xml.Load($Path)
  $id = $xml.OfficeApp.Id
  if ([string]::IsNullOrWhiteSpace($id)) {
    throw "В манифесте $Path не найден элемент Id"
  }
  return $id
}

$targets = @(
  [pscustomobject]@{ Name = 'manifest.xml'; Path = Join-Path $projectRoot 'manifest.xml' }
)
if ($IncludeDev) {
  $targets += [pscustomobject]@{ Name = 'manifest.dev.xml'; Path = Join-Path $projectRoot 'manifest.dev.xml' }
}

if ($Show) {
  Write-Output "Ключ: $developerKey"
  if (-not (Test-Path -LiteralPath $developerKey)) {
    Write-Output '  ключа нет — надстройка не подключена'
    exit 1
  }
  $props = Get-ItemProperty -LiteralPath $developerKey
  $names = $props.PSObject.Properties |
    Where-Object { $_.Name -notlike 'PS*' } |
    Select-Object -ExpandProperty Name
  if (-not $names) {
    Write-Output '  записей нет — надстройка не подключена'
    exit 1
  }
  foreach ($name in $names) {
    $value = $props.$name
    $exists = Test-Path -LiteralPath $value
    $mark = '[ок]'
    if (-not $exists) { $mark = '[нет файла]' }
    Write-Output "  $mark $name -> $value"
  }
  exit 0
}

if ($Remove) {
  if (-not (Test-Path -LiteralPath $developerKey)) {
    Write-Output 'Регистрации нет — снимать нечего.'
    exit 0
  }
  foreach ($target in $targets) {
    if (-not (Test-Path -LiteralPath $target.Path)) { continue }
    $id = Get-ManifestId -Path $target.Path
    Remove-ItemProperty -LiteralPath $developerKey -Name $id -ErrorAction SilentlyContinue
    $subKey = Join-Path $developerKey $id
    if (Test-Path -LiteralPath $subKey) {
      Remove-Item -LiteralPath $subKey -Recurse -Force
    }
    Write-Output "Снята регистрация: $($target.Name) ($id)"
  }
  Write-Output 'Перезапустите Excel, чтобы изменение вступило в силу.'
  exit 0
}

# --- Регистрация ---

if ($ClearCache) {
  $excel = Get-Process EXCEL -ErrorAction SilentlyContinue
  if ($excel) {
    throw 'Excel запущен. Закройте Excel полностью, иначе кэш будет перезаписан обратно при выходе.'
  }
  $wefCache = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
  foreach ($folder in 'AppCommands', 'AddinInfo', 'AggregatedCache') {
    $path = Join-Path $wefCache $folder
    if (Test-Path -LiteralPath $path) {
      Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue
      Write-Output "Кэш очищен: $folder"
    }
  }
  # Признаки готовности кэша ленты: без сброса Excel не перечитывает команды.
  $wefKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF'
  foreach ($name in 'Excel_RibbonCache', 'Excel_AggregatedCache') {
    if ($null -ne (Get-ItemProperty -LiteralPath $wefKey -Name $name -ErrorAction SilentlyContinue)) {
      Set-ItemProperty -LiteralPath $wefKey -Name $name -Value 0
      Write-Output "Сброшен признак кэша: $name"
    }
  }
}

if (-not (Test-Path -LiteralPath $developerKey)) {
  New-Item -Path $developerKey -Force | Out-Null
}

foreach ($target in $targets) {
  if (-not (Test-Path -LiteralPath $target.Path)) {
    throw "Манифест не найден: $($target.Path)"
  }
  $id = Get-ManifestId -Path $target.Path

  # Форма записи взята из резервной копии рабочего состояния
  # (logs/excel-ai-developer-registration-backup.reg): Office читает и значение
  # в самом разделе Developer, и подраздел с путём по умолчанию. Одного
  # значения оказалось недостаточно — Excel его не увидел.
  New-ItemProperty -Path $developerKey -Name $id -Value $target.Path -PropertyType String -Force | Out-Null

  $subKey = Join-Path $developerKey $id
  if (-not (Test-Path -LiteralPath $subKey)) {
    New-Item -Path $subKey -Force | Out-Null
  }
  New-ItemProperty -Path $subKey -Name '(default)' -Value $target.Path -PropertyType String -Force | Out-Null
  # Флаги отладчика выставляем в 0: рабочий режим не должен ждать отладчик.
  foreach ($flag in 'UseDirectDebugger', 'UseWebDebugger', 'UseLiveReload') {
    New-ItemProperty -Path $subKey -Name $flag -Value 0 -PropertyType DWord -Force | Out-Null
  }

  Write-Output "Подключено: $($target.Name)"
  Write-Output "  Id:   $id"
  Write-Output "  Путь: $($target.Path)"
}

Write-Output ''
Write-Output 'Готово. Закройте Excel полностью и откройте снова: реестр читается при запуске.'
Write-Output 'Кнопка появится на вкладке «Главная», группа «AI-панель».'
if (-not $IncludeDev) {
  Write-Output ''
  Write-Output 'Отладочная панель не подключена. Добавить её: scripts\register-addin.ps1 -IncludeDev'
}
