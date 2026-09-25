<#
  Проверка готового комплекта в Windows — то, что увидит пользователь.

  Сервер запускается из комплекта его собственным Node с тем же окружением,
  что выставляет scripts\start-server.ps1. Проверяется: HTTPS доверен (без
  отключения проверки сертификата), панель отдаётся, ключ сохраняется через
  настоящий DPAPI, файл ключей — только шифротекст, после перезапуска ключ
  расшифровывается. Ключ в проверке ненастоящий; в конце проверка убирает
  за собой журналы и файл ключей.

  Нужны свободный порт и доверенный сертификат localhost (в CI ставится
  заранее: office-addin-dev-certs install --machine). Работает и в Windows
  PowerShell 5.1, который есть у любого пользователя, и в PowerShell 7.

  Запуск: pwsh -File scripts/bundle-smoke.ps1 -BundleRoot bundle/ExcelAI
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)] [string] $BundleRoot,
  [int] $Port = 3000
)

$ErrorActionPreference = 'Stop'

$root = (Resolve-Path -LiteralPath $BundleRoot).Path
$node = Join-Path $root 'node\node.exe'
$id = (Get-Content -LiteralPath (Join-Path $root 'releases\current.json') -Raw | ConvertFrom-Json).id
$entry = Join-Path $root "server\releases\$id\dist\server.js"
$logDir = Join-Path $root 'logs'
$keysFile = Join-Path $root 'server\keys.dpapi'
$tokenFile = Join-Path $root 'server\panel-token'
$base = "https://localhost:$Port"
$testKey = 'sk-smoke-0123456789abcdef'
$failures = [System.Collections.Generic.List[string]]::new()
$starts = 0

function Check([string] $Name, [bool] $Ok, [string] $Detail = '') {
  if ($Ok) {
    Write-Output "  [ок]   $Name"
  } else {
    Write-Output "  [нет]  $Name $Detail"
    $failures.Add($Name)
  }
}

function Start-BundleServer {
  $script:starts++
  New-Item -ItemType Directory -Path $logDir -Force | Out-Null
  $env:EXCEL_AI_PROJECT_ROOT = $root
  $env:PANEL_DIST_DIR = Join-Path $root "releases\$id\panel"
  $env:EXCEL_AI_RELEASE_ID = $id
  # Сервер берёт порт из окружения: без этого -Port менял только адрес проверки.
  $env:PORT = "$Port"
  $process = Start-Process -FilePath $node -ArgumentList "`"$entry`"" -WorkingDirectory $root -PassThru -NoNewWindow `
    -RedirectStandardOutput (Join-Path $logDir "smoke-$starts.out.log") `
    -RedirectStandardError (Join-Path $logDir "smoke-$starts.err.log")
  $deadline = (Get-Date).AddSeconds(90)
  while ((Get-Date) -lt $deadline) {
    if ($process.HasExited) { throw "Сервер завершился при запуске с кодом $($process.ExitCode)." }
    try {
      $health = Invoke-RestMethod -Uri "$base/api/health" -TimeoutSec 3
      return [pscustomobject]@{ Process = $process; Health = $health }
    } catch {
      Start-Sleep -Milliseconds 500
    }
  }
  Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
  throw "Сервер не ответил по $base за 90 секунд."
}

function Stop-BundleServer($server) {
  if ($server -and -not $server.Process.HasExited) {
    Stop-Process -Id $server.Process.Id -Force
    $server.Process.WaitForExit()
  }
}

function Get-KeyStatus($state, [string] $ProviderId) {
  return $state.providers | Where-Object { $_.id -eq $ProviderId }
}

# Ключи из окружения раннера не должны подменить проверку ключей из панели.
foreach ($name in @([Environment]::GetEnvironmentVariables().Keys)) {
  if ($name -match '_API_KEY$|^QWEN_BASE_URL$') { Remove-Item -LiteralPath "env:$name" }
}

Write-Output "Проверка комплекта $root, выпуск $id"
$nodeVersion = & $node --version
Check "свой Node запускается ($nodeVersion)" ($LASTEXITCODE -eq 0)
Check 'в комплекте нет каталогов разработки' (-not (Test-Path (Join-Path $root 'dist')) -and
  -not (Test-Path (Join-Path $root 'server\dist')) -and -not (Test-Path (Join-Path $root '.git')))
Check 'в комплекте нет ключей и server\.env' (-not (Test-Path $keysFile) -and -not (Test-Path (Join-Path $root 'server\.env')))
# Токен у каждой установки свой: в архив он попадать не должен.
Check 'в комплекте нет токена панели' (-not (Test-Path $tokenFile))

$server = $null
try {
  $server = Start-BundleServer
  Check 'HTTPS доверен, /api/health отвечает' ($server.Health.app -eq 'excel-ai-addin')
  Check "выбран выпуск $id" ($server.Health.release -eq $id)

  $page = Invoke-WebRequest -UseBasicParsing -Uri "$base/taskpane.html"
  Check 'панель отдаётся' ($page.StatusCode -eq 200 -and $page.Content -match 'id="root"')

  # Без токена API не отвечает (8.0.1); токен сервер создал при запуске.
  $denied = try { Invoke-WebRequest -UseBasicParsing -Uri "$base/api/keys" -TimeoutSec 5 | Out-Null; 0 } catch { [int]$_.Exception.Response.StatusCode }
  Check 'без токена панели API отвечает 401' ($denied -eq 401)
  $auth = @{ 'X-Panel-Token' = ([System.IO.File]::ReadAllText($tokenFile)).Trim() }
  $state = Invoke-RestMethod -Uri "$base/api/keys" -Headers $auth
  Check 'хранение ключей доступно (DPAPI)' ($state.storage.available -eq $true)
  Check 'до сохранения ключей нет' (-not ($state.providers | Where-Object { $_.source }))

  $body = @{ key = $testKey } | ConvertTo-Json
  $state = Invoke-RestMethod -Method Put -Uri "$base/api/keys/deepseek" -ContentType 'application/json' -Body $body -Headers $auth
  $deepseek = Get-KeyStatus $state 'deepseek'
  Check 'ключ сохранён, виден только хвост' ($deepseek.source -eq 'panel' -and $deepseek.hint.EndsWith('cdef') -and
    -not ((ConvertTo-Json $state -Depth 5).Contains($testKey)))
  Check 'файл ключей содержит только шифротекст' ((Test-Path $keysFile) -and
    -not ((Get-Content -LiteralPath $keysFile -Raw).Contains($testKey)))

  $providers = Invoke-RestMethod -Uri "$base/api/providers" -Headers $auth
  Check 'DeepSeek появился среди доступных' ([bool]($providers | Where-Object { $_.id -eq 'deepseek' }))
} catch {
  Check 'без исключений' $false "$($_.Exception.Message)"
} finally {
  Stop-BundleServer $server
}

$server = $null
try {
  $server = Start-BundleServer
  $auth = @{ 'X-Panel-Token' = ([System.IO.File]::ReadAllText($tokenFile)).Trim() }
  $state = Invoke-RestMethod -Uri "$base/api/keys" -Headers $auth
  $deepseek = Get-KeyStatus $state 'deepseek'
  Check 'после перезапуска ключ расшифрован' ($deepseek.source -eq 'panel' -and -not $state.storage.error)

  $state = Invoke-RestMethod -Method Delete -Uri "$base/api/keys/deepseek" -Headers $auth
  Check 'ключ удаляется' (-not (Get-KeyStatus $state 'deepseek').source)
} catch {
  Check 'без исключений' $false "$($_.Exception.Message)"
} finally {
  Stop-BundleServer $server
}

if ($failures.Count) {
  Write-Output ''
  Write-Output "Не прошло проверок: $($failures.Count). Журнал сервера:"
  Get-ChildItem -LiteralPath $logDir -Filter '*.log' | ForEach-Object {
    Write-Output "--- $($_.Name)"
    Get-Content -LiteralPath $_.FullName -Tail 40
  }
  exit 1
}

# Комплект уходит пользователям: следов проверки в нём быть не должно.
Remove-Item -LiteralPath $logDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $keysFile -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $tokenFile -Force -ErrorAction SilentlyContinue
Write-Output 'Комплект исправен.'
