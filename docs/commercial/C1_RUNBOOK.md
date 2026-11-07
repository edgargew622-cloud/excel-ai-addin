# C1 — запуск и эксплуатация закрытого тестового backend

Только локальная разработка на базе 1.0.48. Никаких production аккаунтов, реальных книг, AI-ключей или платёжных данных. [Статус реализации](C1_PROGRESS.md) и [архитектура](SAAS_ARCHITECTURE.md).

## 1. PostgreSQL и проверки

Нужны Node 24 и Docker с Linux engine. PostgreSQL image/digest закреплён в [compose.commercial.yml](../../compose.commercial.yml). Фиксированный пароль в Compose — синтетическая локальная fixture, не production secret. Порт открыт только на loopback.

```powershell
docker compose -f compose.commercial.yml up -d --wait
npm ci --ignore-scripts --no-audit --no-fund
npm --prefix server ci --ignore-scripts --no-audit --no-fund
$env:COMMERCIAL_TEST_DATABASE_URL = 'postgresql://amai_c1:c1-local-test-only@127.0.0.1:55432/amai_c1_test'
npm --prefix server run test:commercial
npm run check
npm run build:web
```

**Тестовый suite очищает только схему `commercial` в специально указанной базе, имя которой заканчивается на `_c1_test`.** Не используйте её для ручной разработки, хранения полезных ключей или данных. Этот guard не означает, что база с именем `_c1_test` автоматически безопасна: оператор обязан указать именно изолированную fixture.

На текущей машине PostgreSQL уже запущен как `amai-c1-postgres`, порт 55432. Не запускайте второй Compose на тот же порт; используйте существующий контейнер либо отдельный тестовый порт. Его данные синтетические. В CI PostgreSQL запускается в отдельном service container.

## 2. Отдельная база разработки

Не смешивать с очищаемой integration database. При использовании Compose:

```powershell
docker compose -f compose.commercial.yml exec postgres createdb -U amai_c1 amai_c1
Copy-Item server/.env.commercial.example server/.env.commercial
```

В `server/.env.commercial` задайте:

- `COMMERCIAL_ENABLED=1`, `COMMERCIAL_AUTH_MODE=test`, `COMMERCIAL_FAKE_PROVIDER=1`.
- `COMMERCIAL_DATABASE_URL` из example для отдельной базы `amai_c1`.
- `COMMERCIAL_DEV_MASTER_KEY`: 32 случайных байта как 64 hex-символа; значение хранить только локально в ignored env-файле. Никогда не использовать реальный KMS/platform secret.
- `COMMERCIAL_DEV_SEED=1` только для явного создания fixture.

Команда генерации локального dev-ключа (не вставляйте результат в issue/чат/коммит):

```powershell
node -e "console.log(require('node:crypto').randomBytes(32).toString('hex'))"
npm --prefix server run commercial:migrate
npm --prefix server run commercial:seed
npm --prefix server run commercial
```

CLI использует `--env-file-if-exists=.env.commercial` относительно `server/`. Seed пишет `server/.commercial-dev.json` с синтетическими org/user/token; stdout не содержит token. Файл ignored, токен истекает через сутки. Повтор при существующем файле запрещён; нельзя повторять seed, чтобы «обнулить» счёт. Seed не сбрасывает глобальный бюджет. После создания fixture верните `COMMERCIAL_DEV_SEED=0`.

Server стартует на `http://127.0.0.1:3300`, не на 3000; локальная Excel-панель и её токены/ключи не меняются. `COMMERCIAL_ENABLED` по умолчанию выключен. `NODE_ENV=production`, public bind и remote DB для включённого C1 отвергаются до подключения.

## 3. Тестовый запрос

В другом PowerShell:

```powershell
$fixture = Get-Content server/.commercial-dev.json -Raw | ConvertFrom-Json
$headers = @{
  Authorization = "Bearer $($fixture.token)"
  'Idempotency-Key' = [guid]::NewGuid().ToString()
}
$body = @{
  provider = 'fake'; model = 'c1-test'; keyMode = 'hosted'
  messages = @(@{role='user'; content='synthetic request'})
} | ConvertTo-Json -Depth 5
Invoke-WebRequest -UseBasicParsing -Method Post `
  -Uri "http://127.0.0.1:3300/api/v1/orgs/$($fixture.orgId)/chat" `
  -Headers $headers -ContentType 'application/json' -Body $body
```

Ответ — SSE `request`, `delta`, `done`. Повтор тех же body/idempotency-key возвращает состояние, не новый текст и не второй вызов модели. Новый body со старым key → 409. Fake quote = 10 micro-units, фактическое тестовое списание = 3; это не реальная цена LLM. BYOK потребляет 1 service_request, отдельно записывает provider cost и не тратит hosted credits.

## 4. API и безопасность

| Endpoint | Назначение |
| --- | --- |
| `GET /api/v1/health`, `/ready` | Минимальная liveness/readiness, без secrets |
| `GET /api/v1/me` | Только активные собственные memberships |
| `GET /api/v1/orgs/:orgId/entitlements` | Server-side subscription projection |
| `POST .../chat` | Strict body, finite quote, transaction reserve, durable dispatch, fake provider, settlement |
| `GET .../usage`, `GET .../usage/:id` | Собственные запросы; owner/admin — запросы своей организации; чужой tenant → 404 |
| `GET/POST .../keys` | Только metadata/hint на чтение; private или явно shared test key |
| `POST .../keys/:id/rotate`, `DELETE .../keys/:id` | Shared key — owner/admin; private — владелец; revoked ciphertext удаляется |

Нельзя передавать `userId`, `role`, `price`, `baseURL` или произвольные tools в chat body. Служебная identity определяется bearer token и текущей БД. Доступ к PostgreSQL и dev master key равнозначен административному доступу к test-контуру; не открывайте их наружу. В частности, reverse proxy не делает test auth production-ready.

## 5. Неопределённый расход / incident

1. `reserved` без dispatch можно безопасно освободить. `dispatched` после сбоя становится `uncertain`; резерв остаётся, автоматического повтора provider нет.
2. Перед повторным пользовательским действием запросите usage по ID или повторите исходный idempotency-key для чтения статуса. Не создавайте новый key, чтобы обойти pending spend.
3. Reconciler проверяет stale records при старте и раз в минуту, стандартный возраст 180 секунд. Не уменьшайте его ниже верхнего времени штатного provider/DB вызова в реальном адаптере.
4. Trusted reconciliation принимает **подтверждённые наблюдения провайдера**, а не цену, присланную пользователем. В C1 это внутренний `Ledger.settle`; публичного HTTP endpoint нет. Uncertain не считается бесплатным.
5. При observed cost выше quote блокируются затронутые global/org/user accounts, сохраняется резерв и отдельная `overrun` запись. Нужны разбор bounds/price version и подтверждённое решение оператора; автоматического unfreeze/списания сверх cap нет. Не «чините» это ручным обнулением reserved.

В техническом логе только requestId/status/latency/error code. Не добавляйте туда exception objects, SQL, headers, body, provider keys или тексты Excel. Сам ledger содержит tenant/user IDs и токены usage, но не prompts; доступ к БД/логам также ограничивается оператором.

## 6. Откат и восстановление

- Откат кода: остановить только коммерческий server, выбрать предыдущий commit. Локальный Excel server/действующие данные книги не затрагивать.
- Схема v1 создаётся транзакционно, checksum защищает от незаметного изменения уже применённой миграции. Не редактировать applied migration; следующие изменения — новая версия.
- `COMMERCIAL_ROLLBACK_EMPTY=1` с `commercial:migrate` удаляет схему только если нет организаций и usage. В populated DB операция откажется. Это **не** механизм production rollback.
- Перед будущей staging миграцией сохранить PostgreSQL backup и отдельно доступ к нужным версиям encryption keys. Восстановление ciphertext без старых ключей не работает; удаление старой keyring version должно быть согласовано с re-encryption/backup retention.
- Успешный тест fresh pool/process crash доказывает сохранение pending state в существующей БД, а не RPO/RTO восстановления PostgreSQL из backup. Restore drill с выбранным hosting/KMS остаётся открытым.

Ничего из этого не публикует релиз, не меняет `production` и не включает реальные расходы. Следующий обязательный шаг для hosted-сервиса — C2 identity и выбранные managed secret/hosting adapters; guards C1 до этого не снимать.
