; Установщик AI-панели для Excel (этап 8, 8.8.4). Собирается в GitHub Actions:
;   ISCC /DAppVersion=1.0.4 /DSourceDir=..\bundle\ExcelAI /DOutputDir=..\bundle packaging\ExcelAI.iss
;
; - Ставит в личную папку пользователя: права администратора не нужны, и
;   другие пользователи компьютера папку не видят.
; - Обновление — тот же Setup поверх: сервер этой папки останавливается до
;   копирования файлов (иначе занятый node.exe не заменить), ключи и токен
;   панели сохраняются.
; - Удаление — через «Приложения» Windows: снимается автозапуск и регистрация
;   в Excel, папка удаляется вместе с ключами.
; - Установщик не подписан: при первом запуске Windows SmartScreen покажет
;   «Неизвестный издатель». Подпись — отдельное решение (платный сертификат).

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#ifndef SourceDir
  #define SourceDir "..\bundle\ExcelAI"
#endif
#ifndef OutputDir
  #define OutputDir "..\bundle"
#endif

[Setup]
AppId={{5F2C9A41-1F7C-4F2B-9A0E-9C1F2D7B3E10}
AppName=AI-панель для Excel
AppVersion={#AppVersion}
AppVerName=AI-панель для Excel {#AppVersion}
AppPublisher=edgargew622-cloud
AppPublisherURL=https://github.com/edgargew622-cloud/excel-ai-addin
AppSupportURL=https://github.com/edgargew622-cloud/excel-ai-addin
AppUpdatesURL=https://github.com/edgargew622-cloud/excel-ai-addin/releases
DefaultDirName={localappdata}\Programs\ExcelAI
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir={#OutputDir}
OutputBaseFilename=ExcelAI-Setup-{#AppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName=AI-панель для Excel
LicenseFile={#SourceDir}\LICENSE
CloseApplications=no

[Languages]
Name: "ru"; MessagesFile: "compiler:Languages\Russian.isl"

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[UninstallDelete]
; Файлы, появившиеся после установки: ключи, токен, журналы, каталог.
Type: filesandordirs; Name: "{app}"

[Code]
function ExcelRunning(): Boolean;
var
  Code: Integer;
begin
  Result := Exec(ExpandConstant('{cmd}'), '/c tasklist /FI "IMAGENAME eq EXCEL.EXE" | find /I "EXCEL.EXE" >nul', '', SW_HIDE, ewWaitUntilTerminated, Code) and (Code = 0);
end;

function RunScript(const Script, Args, LogName: String): Integer;
var
  Code: Integer;
begin
  Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
    '-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "& ''' + ExpandConstant('{app}') + '\scripts\' + Script + ''' ' + Args +
    ' *> ''' + ExpandConstant('{app}') + '\' + LogName + '''"',
    ExpandConstant('{app}'), SW_HIDE, ewWaitUntilTerminated, Code);
  Result := Code;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Code: Integer;
begin
  Result := '';
  while ExcelRunning() do
    if MsgBox('Закройте Excel полностью: пока он открыт, он держит надстройку, и обновление не дойдёт до ленты.' + #13#10#13#10 +
              'Закройте Excel и нажмите «Повтор».', mbError, MB_RETRYCANCEL) = IDCANCEL then
    begin
      Result := 'Установка отменена: Excel открыт.';
      exit;
    end;
  // Обновление: сервер из этой папки держит node.exe — останавливаем его до копирования.
  Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
    '-NoProfile -NonInteractive -Command "Get-CimInstance Win32_Process | Where-Object { ($_.Name -eq ''node.exe'' -or $_.Name -eq ''powershell.exe'') -and $_.CommandLine -and $_.CommandLine.IndexOf(''' +
    ExpandConstant('{app}') + ''', [StringComparison]::OrdinalIgnoreCase) -ge 0 } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"',
    '', SW_HIDE, ewWaitUntilTerminated, Code);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    WizardForm.StatusLabel.Caption := 'Настройка: сертификат, автозапуск, сервер, регистрация в Excel…';
    if RunScript('install.ps1', '', 'install.log') <> 0 then
      MsgBox('Установка файлов прошла, но настройка не завершилась.' + #13#10 +
             'Подробности — в ' + ExpandConstant('{app}') + '\install.log. Можно запустить «Установить.cmd» в этой папке ещё раз.',
             mbError, MB_OK);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Code: Integer;
begin
  if CurUninstallStep = usUninstall then
  begin
    while ExcelRunning() do
      if MsgBox('Закройте Excel, чтобы удалить надстройку полностью, и нажмите «Повтор».', mbError, MB_RETRYCANCEL) = IDCANCEL then
        break;
    Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
      '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + ExpandConstant('{app}') + '\scripts\uninstall.ps1"',
      ExpandConstant('{app}'), SW_HIDE, ewWaitUntilTerminated, Code);
  end;
end;

[Messages]
ru.FinishedLabel=AI-панель установлена.%n%nОткройте Excel: на вкладке «Главная» появится кнопка «Открыть чат». Затем нажмите «Ключи» и вставьте ключ провайдера модели.
