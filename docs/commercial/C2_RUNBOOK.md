# C2 — закрытый локальный пилот: инструкция

Статус: [C2_PROGRESS.md](C2_PROGRESS.md). Базовый SHA `355f2cfa65e15fc6c3d38887db0c65c02d97a861`; version 1.0.48. Работать в `codex/commercial-c2`. Все книги/суммы синтетические. Реальных AI-вызовов этот entrypoint не выполняет.

## 1. Окружение и проверки

Node 24, PostgreSQL 17, Docker Linux engine. На текущем компьютере существующий контейнер `amai-c1-postgres`, loopback 55432; не занимайте тот же порт вторым Compose. Для нового окружения используйте [compose.commercial.yml](../../compose.commercial.yml), затем отдельную dev DB по [C1 runbook](C1_RUNBOOK.md).

```powershell
npm ci --ignore-scripts --no-audit --no-fund
npm --prefix server ci --ignore-scripts --no-audit --no-fund
$env:COMMERCIAL_TEST_DATABASE_URL='postgresql://amai_c1:c1-local-test-only@127.0.0.1:55432/amai_c1_test'
npm --prefix server test
npm --prefix server run test:commercial
npm --prefix server run test:commercial:c2
npm test
npm run build
npm run build:web
$env:VITE_COMMERCIAL_C2='1'
npm run build
Remove-Item Env:VITE_COMMERCIAL_C2
```

Только специально выделенная база `*_c1_test`: suites очищают fixture schema. **Не запускайте их одновременно в одной базе.** Не направляйте их на dev DB `amai_c1`. Никаких customer/Excel данных и полезных ключей в fixture DB.

Каталог обновляется `npm run commercial:catalog` после изменения existing tool registry; root test проверяет drift. Не редактировать generated JSON вручную.

## 2. Microsoft Entra ID

Выбран пользователем. Нужен доступ владельца к тестовому tenant; платные cloud resources не создаются этими командами.

1. Entra admin center → App registrations → New registration. Поддерживаемые accounts: **только этот tenant**.
2. Добавить платформу **Web**, redirect URI ровно `https://localhost:3100/api/v1/auth/callback`. Это серверный confidential-client flow, **не SPA**, implicit flow не включать. Authorization Code + PKCE S256 реализован на сервере.
3. Скопировать Application (client) ID и Directory (tenant) ID. Создать короткоживущий тестовый Client Secret, значение записать **только** в ignored `server/.env.commercial`. Не отправлять его в чат, issue, Git, screenshots или логи. Ограничить доступ к файлу своему Windows account. Секрет предназначен только локальному development, не будущему hosting.
4. Разрешить тестовому пользователю вход согласно tenant policy. Код запрашивает `openid profile`, не запрашивает Graph, refresh token или доступ ко всем книгам.
5. Найти **Object ID пользователя внутри этого tenant**. Для гостя — guest object в этом tenant. `oid` берётся только из проверенного ID token; email/имя/клиентский userId не авторизация. Отсутствующий allowlist даёт отказ `pilot_invitation_required`.

Основание выбора потока: [Office dialog authentication](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/auth-with-office-dialog-api), [Office dialog API](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins), [Entra authorization code + PKCE](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow). Реальная совместимость конкретного Office host пока pending; не меняйте registry/security/privacy settings ради обхода ограничений.

## 3. Серверная конфигурация и приглашение

Скопировать [example](../../server/.env.commercial.example) в `server/.env.commercial`. Не перезаписывать существующий локальный файл. Использовать отдельную dev DB `amai_c1` (не очищаемую suite), случайный 32-byte dev master key, `COMMERCIAL_ENABLED=1`, `COMMERCIAL_AUTH_MODE=test`, `COMMERCIAL_FAKE_PROVIDER=1` и `COMMERCIAL_DEV_SEED=1` для **однократного** создания fixture по C1 runbook. Старый fixture bearer C2 не принимает.

Затем заполнить:

```dotenv
COMMERCIAL_C2_ENABLED=1
COMMERCIAL_BIND=127.0.0.1
COMMERCIAL_PORT=3300
COMMERCIAL_PANEL_ORIGIN=https://localhost:3100
COMMERCIAL_ENTRA_TENANT=<tenant-uuid>
COMMERCIAL_ENTRA_CLIENT_ID=<application-uuid>
COMMERCIAL_ENTRA_CLIENT_SECRET=<server-only-test-secret>
COMMERCIAL_PILOT_PROVISION=1
COMMERCIAL_PILOT_USER_ID=<existing-synthetic-userId-from-local-fixture>
COMMERCIAL_PILOT_OBJECT_ID=<user-object-id-in-selected-tenant>
```

```powershell
npm --prefix server run commercial:migrate
npm --prefix server run commercial:pilot-user
npm --prefix server run commercial:c2
```

После provisioning вернуть `COMMERCIAL_PILOT_PROVISION=0` и `COMMERCIAL_DEV_SEED=0`. Команда изменяет только identity mapping существующего fixture user, не роль/квоты. Новый вход не создаёт пользователей автоматически. Нельзя повторять seed для обхода бюджета. Не запускать C1 server на том же порту одновременно с C2.

Сервер C2 связывается с Entra/JWKS по фиксированным HTTPS URLs; отказ IdP запрещает новый вход. Никакой test identity fallback нет. `NODE_ENV=production`, public bind и remote DB всё ещё отвергаются. Не снимайте эти guards для публичной демонстрации.

## 4. Панель и сессия

В отдельном PowerShell:

```powershell
$env:VITE_COMMERCIAL_C2='1'
npm run dev
```

Dev taskpane — `https://localhost:3100/taskpane.html`, gateway `/api/v1` проксируется на `http://127.0.0.1:3300` с сохранением исходного Origin. `/api` локального режима остаётся прежним. Установленные manifests/релиз 1.0.48 автоматически не переключаются на эту dev-панель. Подключить именно dev manifest через существующий development/sideload workflow; проверить его SourceLocation/port прежде чем тестировать. Не устанавливать новые сертификаты без отдельного решения владельца, не обходить предупреждения недоверенного сертификата.

Существующий [manifest.dev.xml](../../manifest.dev.xml) уже указывает на порт 3100. После запуска Vite и при доверенном существующем dev-сертификате команда `npx office-addin-debugging start manifest.dev.xml` подключает dev-панель; выполнять её на синтетической книге в согласованной тестовой Excel-сессии. Это действие не выполнено автоматически в текущей поставке.

Включить checkbox **«Закрытый пилот am.AI»**, нажать «Войти через Microsoft», выбрать свою организацию. Build flag и checkbox по умолчанию выключены. Кнопка «Ключи» локального режима в пилоте не нужна. Память, локальные вложения и интернет не отправляются коммерческим transport автоматически.

Office Dialog должен поддерживать сообщения с проверяемым `origin`; неизвестный origin отклоняется. Callback отправляет только ticket на конкретный origin; панель подтверждает владение начатым flow через parent proof. App bearer в памяти, срок 1 час; refresh токен не хранится. После reload панели нужен новый Dialog; Entra может использовать свою SSO-сессию. Никакой обещанной бесшумной авторизации до host smoke.

Logout отзывает серверную app session и очищает память панели. Это не глобальный выход из Microsoft. При сетевом отказе серверный logout не подтверждён, предупреждение видно; session истечёт не позднее часа. Полезная локальная отмена не зависит от login/subscription.

## 5. Проверка в Excel Windows

**Пока не выполнена.** Не засчитывать unit/integration/browser mock как Excel E2E. Установленная рабочая надстройка не обновляется этим этапом автоматически.

1. Новая пустая синтетическая книга. Создать листы `PaymentsA`, `PaymentsB`; перенести [PaymentsA.csv](fixtures/PaymentsA.csv) и [PaymentsB.csv](fixtures/PaymentsB.csv) в соответствующие диапазоны **A1:D4**. Разделить CSV на четыре столбца, даты/числа проверить в Excel. Изначально листа `PilotResult` быть не должно. Сохранить контрольную копию. Записать SHA (`git rev-parse HEAD`), build flag, Excel version/build и хэши fixtures (`Get-FileHash ... -Algorithm SHA256`).
2. Войти через Microsoft **без AI-ключа пользователя**. Включить ручные подтверждения и снять «Только анализ».
3. Запрос: «Сверь PaymentsA и PaymentsB, диапазоны A1:D4. Создай отдельный лист PilotResult с расхождениями». Fake provider работает только с этой fixture, не понимает произвольные задачи и не считается реальной моделью.
4. Проверить предпросмотр существующего `reconcile_ranges`, подтвердить. Источники не изменились; новый результат содержит совпавший A100, различие суммы A101 и отдельные A102/A103. Точный формат/категории задаёт existing reconciliation engine. Проверить read-back, итог операции и локальную отмену.
5. «Состояние задания»: task ID, число AI requests, конечные states, spent/reserved. Каждое fake обращение: верхняя reservation 10, actual 3 synthetic micro-units. Вызов Excel-инструмента не отдельный AI-запрос; следующий модельный ход с результатом инструмента — отдельный request этого же task. Возможен предварительный `enable_tools`, поэтому число шагов зависит от подключённых групп.
6. Отдельная новая fixture: отказаться от подтверждения; источники/результат не должны измениться. «Только анализ» также не должен допустить запись. Затем отдельно проверить отмену во время ожидания и local undo после logout/отзыва подписки.
7. На тестовом сценарии с исчерпанным бюджетом убедиться, что новый provider attempt отсутствует, в панели видна причина. Для role/subscription/foreign tenant/100 duplicate/concurrency использовать автоматические suites; при ручной инъекции работать только с synthetic DB, не менять фактическую платную политику.
8. Обрыв/закрытие панели: вновь открыть, войти и нажать «Состояние задания». Проверить статус, не нажимать «продолжай» автоматически и не повторять запись. Completed response не переигрывается в Excel. Unknown/dispatched spend сохраняет резерв до trusted reconciliation по C1 runbook.
9. Записать expected/actual, affected/unaffected ranges, подтверждение, undo и состояния requests в локальный отчёт без tokens/ячейковых данных клиентов. Любая неизвестная запись требует проверки книги, не автоматического retry.

## 6. Реальная модель и секреты будущей среды

[DeepSeekPilot](../../server/src/commercial/deepseekPilot.ts) — подготовленный кандидат, контракт проверен имитатором; в `c2Server.ts` не зарегистрирован. Нет скрытого SDK retry или fallback. URL фиксирован, custom URL клиента не принимается. Цены не зашиты как утверждённые: нужны model ID, version/validUntil и input/output micro-USD per million tokens с проверкой актуальной [официальной документации](https://api-docs.deepseek.com/api/create-chat-completion/). Проверить byte/token upper bound для окончательно выбранной модели.

До подключения владелец отдельно утверждает модель, конечный лимит расходов и duration, platform/org/user/task caps, BYOK policy и обработку unknown. После этого зарегистрировать adapter, передав HostedSecretSource; tenant BYOK приходит из CredentialStore. В текущей работе ни real key, ни live request не использовались.

Для будущего hosting выбрать managed secret store/KMS и workload identity: хранить Entra client secret и platform model key в vault, выдать приложению только read конкретных secret versions, запретить exports/logging, настроить rotation/revoke и alerts. SecretProtector должен шифровать с tenant AAD и удерживать нужные key versions для восстановления backup. Не переносить dev master key/env-файл в hosting. Точный SDK/IAM adapter зависит от утверждённого hosting; он не реализован и не подменяется локальным AES protector.

До выхода за localhost: HTTPS origin + exact redirect/origin policy, managed secrets, backup/restore drill, повтор негативных auth/tenant/usage tests, затем реальный Windows smoke. Production deployment, release, merge и платежи остаются запрещены.
