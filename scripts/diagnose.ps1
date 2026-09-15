<#
  Локальная диагностика надстройки.

  Нужна именно отдельным сценарием: если недоступен сам сервер страницы, панель
  не может нарисовать собственную ошибку — Excel показывает пустой прямоугольник.
  Этот сценарий отвечает на вопрос «что именно не отвечает».
#>

[CmdletBinding()]
param([int] $Port = 3000, [int] $DevPort = 3100)

$ErrorActionPreference = 'Continue'

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$appId = 'excel-ai-addin'
$pointerPath = Join-Path $projectRoot 'releases\current.json'
$problems = New-Object System.Collections.Generic.List[string]

function Show-Section { param([string] $Title) Write-Output ''; Write-Output "== $Title" }
function Show-Ok { param([string] $Text) Write-Output "  [ок]   $Text" }
function Show-Bad {
  param([string] $Text)
  Write-Output "  [нет]  $Text"
  $problems.Add($Text)
}

Write-Output "Диагностика AI-панели для Excel"
Write-Output "Проект: $projectRoot"

Show-Section 'Сборка'
$panelEntry = Join-Path $projectRoot 'dist\taskpane.html'
$serverEntry = Join-Path $projectRoot 'server\dist\server.js'
if (Test-Path -LiteralPath $panelEntry) {
  Show-Ok "Панель собрана: $panelEntry"
} else {
  Show-Bad "Панель не собрана. Выполните: npm run build:all"
}
if (Test-Path -LiteralPath $serverEntry) {
  Show-Ok "Сервер собран: $serverEntry"
} else {
  Show-Bad "Сервер не собран. Выполните: npm run build:all"
}
$releaseId = $null
if (Test-Path -LiteralPath $pointerPath) {
  try {
    $selected = Get-Content -LiteralPath $pointerPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $releaseId = [string]$selected.id
    if ($releaseId -notmatch '^[0-9]{8}T[0-9]{6}Z-[a-f0-9]{8}$' -or
        -not (Test-Path -LiteralPath (Join-Path $projectRoot "releases\$releaseId\panel\taskpane.html")) -or
        -not (Test-Path -LiteralPath (Join-Path $projectRoot "server\releases\$releaseId\dist\server.js"))) {
      Show-Bad 'Выбранный выпуск неполон или имеет неверный идентификатор'
      $releaseId = $null
    } else {
      Show-Ok "Выбран рабочий выпуск $releaseId"
    }
  } catch { Show-Bad "Не удалось прочитать текущий выпуск: $($_.Exception.Message)" }
} else {
  Show-Bad 'Рабочий выпуск не выбран. Выполните npm run check, затем npm run release'
}

Show-Section 'Конфигурация'
$envPath = Join-Path $projectRoot 'server\.env'
if (Test-Path -LiteralPath $envPath) {
  # Читаем только имена переменных. Значения ключей не выводятся никогда.
  $names = Select-String -LiteralPath $envPath -Pattern '^\s*([A-Z_]+)\s*=\s*\S' -AllMatches |
    ForEach-Object { $_.Matches[0].Groups[1].Value }
  if ($names) {
    Show-Ok ("server/.env найден, заполнены: " + ($names -join ', '))
  } else {
    Show-Bad 'server/.env найден, но ни одна переменная не заполнена'
  }
} else {
  Show-Bad "server/.env отсутствует. Скопируйте .env.example и впишите ключи"
}

Show-Section 'Порт и процесс'
$listening = $null
try {
  $listening = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction Stop
} catch {
  $listening = $null
}
if ($listening) {
  $owners = $listening | Select-Object -ExpandProperty OwningProcess -Unique
  foreach ($pidValue in $owners) {
    $proc = Get-Process -Id $pidValue -ErrorAction SilentlyContinue
    if ($proc) {
      Show-Ok "Порт $Port слушает процесс $($proc.ProcessName), PID $pidValue"
    } else {
      Show-Ok "Порт $Port слушает PID $pidValue"
    }
  }
} else {
  Show-Bad "Порт $Port никто не слушает. Запустите: scripts\start-server.ps1"
}

Show-Section 'HTTPS и ответ сервера'
$health = $null
try {
  $health = Invoke-RestMethod -Uri "https://127.0.0.1:$Port/api/health" -TimeoutSec 5
  if ($health.app -eq $appId) {
    Show-Ok "Это наш сервер: версия сборки $($health.version), PID $($health.pid), запущен $($health.startedAt)"
    if ($releaseId -and $health.release -ne $releaseId) {
      Show-Bad "Сервер использует выпуск $($health.release), а для следующего запуска выбран $releaseId. Перезапустите сервер после завершения текущих запросов"
    }
  } else {
    Show-Bad "Порт $Port отвечает, но это другая программа. Не завершайте её — освободите порт или измените PORT"
  }
} catch {
  Show-Bad "Сервер не ответил на https://127.0.0.1:$Port/api/health : $($_.Exception.Message)"
}

if ($health -and $health.app -eq $appId) {
  try {
    $providers = Invoke-RestMethod -Uri "https://127.0.0.1:$Port/api/providers" -TimeoutSec 5
    if ($providers -and $providers.Count -gt 0) {
      Show-Ok ("Провайдеры с ключами: " + (($providers | ForEach-Object { $_.id }) -join ', '))
    } else {
      Show-Bad 'Ни одного провайдера с ключом. Заполните server/.env и перезапустите сервер'
    }
  } catch {
    Show-Bad "Список провайдеров недоступен: $($_.Exception.Message)"
  }

  try {
    $panel = Invoke-WebRequest -Uri "https://127.0.0.1:$Port/taskpane.html" -TimeoutSec 5 -UseBasicParsing
    if ($panel.StatusCode -eq 200) {
      Show-Ok 'Страница панели отдаётся сервером'
    } else {
      Show-Bad "Страница панели вернула код $($panel.StatusCode)"
    }
  } catch {
    Show-Bad "Страница панели недоступна: $($_.Exception.Message)"
  }
}

Show-Section 'Сертификат разработки'
# Панель грузится в webview Excel, который молча отказывается открывать
# недоверенный HTTPS — вместо ошибки виден пустой прямоугольник.
#
# Искать в хранилище сертификат с именем localhost бесполезно: лист выписан на
# CN=127.0.0.1, а в доверенные корневые ставится центр сертификации
# «Developer CA for Microsoft Office Add-ins». Поэтому проверяем три вещи:
# наличие центра, срок годности листа из файла и, главное, фактическое
# доверие — запрос с включённой проверкой сертификата.
$caSubject = 'Developer CA for Microsoft Office Add-ins'
$ca = Get-ChildItem -Path Cert:\CurrentUser\Root, Cert:\LocalMachine\Root -ErrorAction SilentlyContinue |
  Where-Object { $_.Subject -like "*$caSubject*" }
if ($ca) {
  Show-Ok 'Центр сертификации Office Add-ins установлен в доверенные корневые'
} else {
  Show-Bad "Центр сертификации '$caSubject' не найден в доверенных корневых. Выполните: npm run certs"
}

$certDir = Join-Path $env:USERPROFILE '.office-addin-dev-certs'
$leafPath = Join-Path $certDir 'localhost.crt'
if (Test-Path -LiteralPath $leafPath) {
  try {
    $leaf = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $leafPath
    $daysLeft = [int]($leaf.NotAfter - (Get-Date)).TotalDays
    $until = $leaf.NotAfter.ToString('yyyy-MM-dd')
    if ($daysLeft -lt 0) {
      Show-Bad "Сертификат истёк $until. Выполните: npm run certs"
    } elseif ($daysLeft -lt 7) {
      Show-Bad "Сертификат истекает через $daysLeft дн. ($until). Продлите заранее: npm run certs"
    } else {
      Show-Ok "Сертификат действует ещё $daysLeft дн. (до $until)"
    }
  } catch {
    Show-Bad "Не удалось прочитать $leafPath : $($_.Exception.Message)"
  }
} else {
  Show-Bad "Файл сертификата не найден: $leafPath. Выполните: npm run certs"
}

# Решающая проверка: запрос без отключения проверки сертификата. Если он
# проходит, цепочка доверена по-настоящему, а не по совпадению имён.
if ($listening) {
  try {
    Invoke-WebRequest -Uri "https://localhost:$Port/api/health" -TimeoutSec 5 -UseBasicParsing | Out-Null
    Show-Ok 'HTTPS доверен фактически: запрос с проверкой сертификата прошёл'
  } catch {
    Show-Bad "HTTPS не доверен: $($_.Exception.Message). Панель в Excel покажет пустой прямоугольник. Выполните: npm run certs"
  }
} else {
  Write-Output '  [инфо] Фактическую проверку доверия пропускаем: сервер не запущен'
}

Show-Section 'Подключение к Excel'
# Самая частая причина «в Excel ничего нет»: манифест лежит на диске, но Excel
# о нём не знает. Проверять наличие файла бессмысленно — нужна запись в реестре.
$developerKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$registered = @()
if (Test-Path -LiteralPath $developerKey) {
  $props = Get-ItemProperty -LiteralPath $developerKey
  $registered = $props.PSObject.Properties |
    Where-Object { $_.Name -notlike 'PS*' -and $_.Name -ne '(default)' }
}
if ($registered) {
  foreach ($entry in $registered) {
    if (Test-Path -LiteralPath $entry.Value) {
      Show-Ok "Подключён манифест: $($entry.Value)"
    } else {
      Show-Bad "Зарегистрирован путь, которого нет: $($entry.Value). Проверьте регистрацию манифеста в Excel"
    }
  }
} else {
  Write-Output '  [инфо] Developer-регистрации нет; проверяем локальный каталог ниже'
}

$catalogKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs'
$catalogs = @()
if (Test-Path -LiteralPath $catalogKey) {
  $catalogs = Get-ChildItem -LiteralPath $catalogKey -ErrorAction SilentlyContinue
}
if ($catalogs) {
  foreach ($item in $catalogs) {
    $url = (Get-ItemProperty -LiteralPath $item.PSPath).Url
    Write-Output "  [инфо] Каталог надёжных надстроек: $url"
  }
} else {
  Write-Output '  [инфо] Каталог надёжных надстроек не зарегистрирован. Это нормально: основной способ — developer-регистрация выше'
}

$catalogManifest = Join-Path $projectRoot 'catalog\manifest.xml'
if (Test-Path -LiteralPath $catalogManifest) {
  $sourceManifest = Join-Path $projectRoot 'manifest.xml'
  if ((Get-FileHash -LiteralPath $sourceManifest -Algorithm SHA256).Hash -eq
      (Get-FileHash -LiteralPath $catalogManifest -Algorithm SHA256).Hash) {
    Show-Ok 'Манифест в локальном каталоге совпадает с рабочим манифестом'
  } else {
    Show-Bad 'Манифест в каталоге устарел. Выполните scripts\register-local-catalog.ps1 и перезапустите Excel'
  }
} else {
  Show-Bad 'В локальном каталоге нет рабочего манифеста. Выполните scripts\register-local-catalog.ps1'
}

Show-Section 'Разработка'
$devListening = $null
try {
  $devListening = Get-NetTCPConnection -LocalPort $DevPort -State Listen -ErrorAction Stop
} catch {
  $devListening = $null
}
if ($devListening) {
  Write-Output "  [инфо] Порт $DevPort слушается: запущен dev-сервер Vite"
} else {
  Write-Output "  [инфо] Порт $DevPort свободен: dev-сервер не запущен, это нормально для обычной работы"
}

Write-Output ''
if ($problems.Count -eq 0) {
  Write-Output 'Итог: проблем не обнаружено.'
  exit 0
}
Write-Output "Итог: проблем — $($problems.Count)."
foreach ($item in $problems) { Write-Output "  - $item" }
exit 1
