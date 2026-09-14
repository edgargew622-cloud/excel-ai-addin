<#
  Полная переустановка подключения надстройки к Excel.

  Зачем отдельный сценарий: Excel кэширует команды ленты и сведения о
  надстройках между запусками. Пока кэш не сброшен, он может не перечитывать
  developer-регистрацию вовсе — журнал загрузки при этом остаётся пустым, то
  есть попытки загрузить надстройку не происходит.

  Сбросить кэш можно только при закрытом Excel: иначе он перезапишет его
  обратно при выходе. Поэтому сценарий ждёт закрытия Excel сам.

  Порядок: дождаться закрытия Excel, сбросить кэш, записать регистрацию,
  при желании запустить Excel обратно.

  Запуск:  powershell -ExecutionPolicy Bypass -File scripts\reinstall-addin.ps1
  С автозапуском Excel:  ... -StartExcel
#>

[CmdletBinding()]
param(
  [switch] $IncludeDev,
  [switch] $StartExcel,
  [int] $WaitSeconds = 180
)

$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$developerKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$wefKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF'
$wefCache = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'

function Get-ManifestId {
  param([string] $Path)
  $xml = New-Object System.Xml.XmlDocument
  $xml.Load($Path)
  $id = $xml.OfficeApp.Id
  if ([string]::IsNullOrWhiteSpace($id)) { throw "В манифесте $Path нет элемента Id" }
  return $id
}

# --- 1. Дождаться закрытия Excel -------------------------------------------

$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excel) {
  Write-Output ''
  Write-Output '  ЗАКРОЙТЕ EXCEL.'
  Write-Output '  Сохраните работу и закройте все окна Excel.'
  Write-Output "  Жду до $WaitSeconds секунд, затем прекращу."
  Write-Output ''

  $waited = 0
  while ($waited -lt $WaitSeconds) {
    Start-Sleep -Seconds 2
    $waited += 2
    $excel = Get-Process EXCEL -ErrorAction SilentlyContinue
    if (-not $excel) { break }
    if ($waited % 10 -eq 0) {
      Write-Output "  ...Excel ещё запущен ($waited с)"
    }
  }

  if ($excel) {
    Write-Output ''
    Write-Output 'Excel так и не закрылся. Ничего не менял: сброс кэша при работающем Excel бесполезен.'
    Write-Output 'Закройте Excel и запустите сценарий снова.'
    exit 2
  }
  Write-Output '  Excel закрыт, продолжаю.'
  # Excel дописывает кэш при выходе не мгновенно.
  Start-Sleep -Seconds 3
}

# --- 2. Сбросить кэш --------------------------------------------------------

Write-Output ''
Write-Output '== Сброс кэша надстроек'
foreach ($folder in 'AppCommands', 'AddinInfo', 'AggregatedCache') {
  $path = Join-Path $wefCache $folder
  if (Test-Path -LiteralPath $path) {
    try {
      Remove-Item -LiteralPath $path -Recurse -Force
      Write-Output "  удалён кэш: $folder"
    } catch {
      Write-Output "  не удалось удалить $folder : $($_.Exception.Message)"
    }
  } else {
    Write-Output "  кэша нет: $folder"
  }
}

# Признаки готовности кэша: без сброса Excel не перечитывает команды ленты.
foreach ($name in 'Excel_RibbonCache', 'Excel_AggregatedCache', 'ExcelOMEXRefreshPending') {
  $existing = Get-ItemProperty -LiteralPath $wefKey -Name $name -ErrorAction SilentlyContinue
  if ($null -ne $existing) {
    Set-ItemProperty -LiteralPath $wefKey -Name $name -Value 0
    Write-Output "  сброшен признак: $name"
  }
}

# Отметка устаревания настройки ленты для текущей локали.
$ribbonExpiry = Get-Item -LiteralPath $wefKey | Select-Object -ExpandProperty Property |
  Where-Object { $_ -like 'Excel_*_RibbonCustomizationExpire' }
foreach ($name in $ribbonExpiry) {
  Remove-ItemProperty -LiteralPath $wefKey -Name $name -ErrorAction SilentlyContinue
  Write-Output "  удалена отметка устаревания ленты: $name"
}

# --- 3. Записать регистрацию ------------------------------------------------

Write-Output ''
Write-Output '== Регистрация надстройки'

$targets = @(Join-Path $projectRoot 'manifest.xml')
if ($IncludeDev) { $targets += Join-Path $projectRoot 'manifest.dev.xml' }

if (-not (Test-Path -LiteralPath $developerKey)) {
  New-Item -Path $developerKey -Force | Out-Null
}

foreach ($manifestPath in $targets) {
  if (-not (Test-Path -LiteralPath $manifestPath)) {
    throw "Манифест не найден: $manifestPath"
  }
  $id = Get-ManifestId -Path $manifestPath

  New-ItemProperty -Path $developerKey -Name $id -Value $manifestPath -PropertyType String -Force | Out-Null

  $subKey = Join-Path $developerKey $id
  if (-not (Test-Path -LiteralPath $subKey)) { New-Item -Path $subKey -Force | Out-Null }
  New-ItemProperty -Path $subKey -Name '(default)' -Value $manifestPath -PropertyType String -Force | Out-Null
  foreach ($flag in 'UseDirectDebugger', 'UseWebDebugger', 'UseLiveReload') {
    New-ItemProperty -Path $subKey -Name $flag -Value 0 -PropertyType DWord -Force | Out-Null
  }
  Write-Output "  подключён: $(Split-Path $manifestPath -Leaf) ($id)"
}

# --- 4. Проверить, что сервер отдаёт панель ---------------------------------

Write-Output ''
Write-Output '== Сервер панели'
try {
  $health = Invoke-RestMethod -Uri 'https://localhost:3000/api/health' -TimeoutSec 5
  Write-Output "  отвечает: $($health.app), сборка $($health.version)"
} catch {
  Write-Output "  НЕ отвечает: $($_.Exception.Message)"
  Write-Output '  Панель откроется пустой. Запустите: Start-ScheduledTask -TaskName ExcelAiAddinServer'
}

# --- 5. Запустить Excel -----------------------------------------------------

Write-Output ''
if ($StartExcel) {
  Write-Output 'Запускаю Excel...'
  Start-Process -FilePath 'excel.exe'
  Write-Output 'Кнопка должна быть на вкладке «Главная», группа «AI-панель».'
} else {
  Write-Output 'Готово. Откройте Excel.'
  Write-Output 'Кнопка должна быть на вкладке «Главная», группа «AI-панель».'
}
Write-Output ''
Write-Output 'Если кнопки нет, пришлите вывод:'
Write-Output '  Get-Content "$env:TEMP\OfficeAddins.log.txt" -Tail 30'
