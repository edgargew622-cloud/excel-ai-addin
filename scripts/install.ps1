<#
  Установка AI-панели для Excel из готового комплекта. Обычно её запускают
  двойным щелчком по «Установить.cmd» в корне комплекта; вручную:

    powershell -ExecutionPolicy Bypass -File scripts\install.ps1

  Шаги:
    1. сертификат для https://localhost — Windows спросит, доверять ли ему;
    2. автозапуск сервера при входе в Windows;
    3. запуск сервера сейчас и проверка, что отвечает именно этот выпуск;
    4. регистрация надстройки в Excel (общая папка надёжных надстроек).

  Повторный запуск безопасен: каждый шаг либо уже сделан, либо делается заново.
  Ключи провайдеров вводятся потом в самой панели кнопкой «Ключи».
#>

[CmdletBinding()]
param(
  # В CI сертификат ставится для всей машины заранее, без окна подтверждения.
  [switch] $SkipCertificate,
  # Запуск двойным щелчком по «Установить.cmd»: окно ждёт Enter, чтобы
  # человек успел прочитать итог, и ошибка не пропадает вместе с окном.
  [switch] $Pause,
  [string] $TaskName = 'ExcelAiAddinServer'
)

# Манифест надстройки указывает на https://localhost:3000 — другой порт Excel не найдёт.
$Port = 3000

$ErrorActionPreference = 'Stop'

$exitCode = 0
try {
  $root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path

  # Двойной щелчок по «Установить.cmd» прямо внутри архива: Проводник тайком
  # распаковывает файл во временную папку и запускает оттуда. Надстройка
  # «установилась» бы из места, которое Windows скоро сотрёт.
  $temp = [System.IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
  if (($root + '\').StartsWith($temp, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Установка запущена прямо из архива, без распаковки. Закройте это окно, щёлкните по архиву правой кнопкой → «Извлечь все…» → впишите C:\ → «Извлечь», затем дважды щёлкните «Установить.cmd» в папке C:\ExcelAI."
  }
  $downloads = Join-Path $env:USERPROFILE 'Downloads'
  if (($root + '\').StartsWith($downloads + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
    Write-Output "Внимание: надстройка будет работать из папки «Загрузки» ($root). Если потом её удалить или почистить «Загрузки», надстройка перестанет работать. Надёжнее распаковать в C:\ — см. INSTALL.md."
  }
  $bundledNode = Join-Path $root 'node\node.exe'
  $node = if (Test-Path -LiteralPath $bundledNode) { $bundledNode } else { 'node.exe' }
  $pointer = Join-Path $root 'releases\current.json'
  if (-not (Test-Path -LiteralPath $pointer)) {
    throw "Не найден $pointer. Запускайте установку из распакованного комплекта, а в рабочей копии сначала выполните npm run release."
  }
  $release = (Get-Content -LiteralPath $pointer -Raw | ConvertFrom-Json).id

  function Get-Health {
    try {
      return Invoke-RestMethod -Uri "https://localhost:$Port/api/health" -TimeoutSec 3
    } catch {
      return $null
    }
  }

  function Wait-Health([int] $Seconds) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
      $health = Get-Health
      if ($health) { return $health }
      Start-Sleep -Seconds 1
    }
    return $null
  }

  # Файлы из скачанного архива помечены «из интернета». PowerShell с Bypass
  # их и так выполнит, но снятая метка избавляет от лишних предупреждений
  # Windows при следующих запусках.
  Get-ChildItem -LiteralPath $root -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

  Write-Output "Установка AI-панели из $root (выпуск $release)"

  Write-Output '1/6 Доступ к папке надстройки — только у вас'
  # Папка в C:\ наследует право изменения для всех пользователей компьютера:
  # другой пользователь Windows мог бы подменить start-server.ps1, и он
  # запустился бы от вашего имени при входе (найдено 26 сентября 2026 года).
  # Оставляем текущего пользователя, систему и администраторов; SID вместо
  # имён — имена групп на русской Windows другие.
  $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
  & icacls.exe $root /inheritance:r /grant:r "*${sid}:(OI)(CI)F" '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' /C /Q | Out-Null
  if ($LASTEXITCODE -ne 0) { throw "Не удалось закрыть доступ к папке $root (icacls, код $LASTEXITCODE)." }
  $foreign = @((Get-Acl -LiteralPath $root).Access | ForEach-Object { $_.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value } |
    Where-Object { $_ -notin @($sid, 'S-1-5-18', 'S-1-5-32-544') })
  if ($foreign.Count) { throw "У папки $root остались чужие права: $($foreign -join ', ')." }
  Write-Output '     изменять файлы надстройки можете только вы и администраторы компьютера'

  Write-Output '2/6 Сертификат для https://localhost'
  if ($SkipCertificate) {
    Write-Output '     пропущен по параметру -SkipCertificate'
  } else {
    $cli = Join-Path $root 'server\node_modules\office-addin-dev-certs\lib\cli.js'
    if (-not (Test-Path -LiteralPath $cli)) { throw "Не найден установщик сертификата: $cli" }
    Write-Output '     Если Windows спросит, установить ли сертификат, ответьте «Да»: без него Excel покажет пустую панель.'
    & $node $cli install
    if ($LASTEXITCODE -ne 0) { throw 'Сертификат не установлен. Запустите установку ещё раз и подтвердите установку сертификата.' }
  }

  Write-Output '3/6 Автозапуск сервера при входе в Windows'
  & (Join-Path $PSScriptRoot 'register-autostart.ps1') -TaskName $TaskName

  Write-Output '4/6 Запуск сервера'
  $health = Get-Health
  if (-not $health) {
    Start-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    $health = Wait-Health 20
  }
  if (-not $health) {
    # Задача планировщика может не стартовать вне интерактивного входа (так в
    # CI). Супервизор держит один экземпляр через именованный мьютекс, поэтому
    # прямой запуск не создаст второго сервера.
    Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList @(
      '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', "`"$(Join-Path $PSScriptRoot 'start-server.ps1')`""
    ) | Out-Null
    $health = Wait-Health 60
  }
  if (-not $health) {
    throw "Сервер не ответил на https://localhost:$Port. Подробности — в $(Join-Path $root 'logs\server.log') и $(Join-Path $root 'logs\app.log')."
  }
  if ($health.app -ne 'excel-ai-addin') {
    throw "Порт $Port занят другой программой. Освободите его и запустите установку снова."
  }
  if ($health.release -ne $release) {
    throw "На порту $Port уже работает другая копия надстройки (выпуск $($health.release)). Остановите её и запустите установку снова."
  }
  Write-Output "     сервер отвечает, выпуск $($health.release)"

  Write-Output '5/6 Регистрация надстройки в Excel'
  # Без прав администратора (8.8.2): запись в раздел текущего пользователя.
  & (Join-Path $PSScriptRoot 'register-addin.ps1')

  Write-Output '6/6 Обновление надстройки в Excel'
  # Excel держит кнопки ленты и манифест в своём кэше и при обновлении
  # продолжал открывать панель по старому адресу (проверено 26 сентября
  # 2026 года). Удаляем из кэша только записи этой надстройки.
  $addinId = [regex]::Match([System.IO.File]::ReadAllText((Join-Path $root 'manifest.xml')), '<Id>([^<]+)</Id>').Groups[1].Value
  if (-not $addinId) { throw 'В manifest.xml не найден Id надстройки.' }
  if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    Write-Output '     Excel открыт. Чтобы он увидел эту версию, закройте его полностью и запустите установку ещё раз.'
  } else {
    $wef = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
    $stale = @(Get-ChildItem -LiteralPath $wef -Recurse -File -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -like "$addinId*" -and $_.FullName -notlike '*\webview2\*' })
    $stale | Remove-Item -Force -ErrorAction SilentlyContinue
    Write-Output "     кэш надстройки в Office сброшен (записей: $($stale.Count)); Excel прочитает новую версию при запуске"
  }

  Write-Output ''
  Write-Output 'Готово. Дальше в Excel:'
  Write-Output '  1. Полностью закройте и снова откройте Excel.'
  Write-Output '  2. Главная → Надстройки → Мои надстройки → Общая папка → AI-панель → Добавить.'
  Write-Output '  3. Нажмите «Открыть чат», затем «Ключи» — и вставьте ключ провайдера.'
} catch {
  $exitCode = 1
  Write-Host ''
  Write-Host "Установка не завершена: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host 'Исправьте причину и запустите установку ещё раз: повторный запуск безопасен.'
  if (-not $Pause) { throw }
} finally {
  if ($Pause) {
    Write-Host ''
    [void](Read-Host 'Нажмите Enter, чтобы закрыть окно')
  }
}
exit $exitCode
