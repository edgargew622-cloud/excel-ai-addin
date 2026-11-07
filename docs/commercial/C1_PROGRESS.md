# C1 — коммерческий backend: реализация и проверка

Дата: **8 ноября 2026, Europe/Moscow**. Ветка: `codex/commercial-c1`. База кода — опубликованный `v1.0.48`, SHA `23830108707dbb1856537b1e22798cb25b164631`. Архитектура C0 принята пользователем для начала C1; это не разрешение на merge/deployment или реальные расходы.

## Объём текущей поставки

Реализован закрытый локальный контур C1: PostgreSQL, пользователи/организации/membership, server-side entitlement, атомарные резервы бюджета, durable attempt/usage ledger, BYOK service quota, tenant-scoped encrypted credential store и отдельный HTTP entrypoint с тестовым провайдером. Он не подключён к текущему интерфейсу Excel и не заменяет локальный `/api/chat`.

Текущий статус — **C1 foundation реализован; полный допуск к hosted/staging-продукту ещё не выдан**. C2 identity, управляемый KMS, выбранный staging hosting и реальные provider calls остаются следующими интеграциями. Нельзя называть эту сборку production SaaS или считать fake-тесты проверкой всех реальных моделей.

## Что добавлено

| Область | Реализация | Проверка |
| --- | --- | --- |
| Отдельный сервер | [server.ts](../../server/src/commercial/server.ts), [app.ts](../../server/src/commercial/app.ts), [core.ts](../../server/src/commercial/core.ts) | По умолчанию выключен. Включённый C1 отвергает production, public bind и remote database; только loopback; paid providers отсутствуют |
| Identity / tenant | [identity.ts](../../server/src/commercial/identity.ts), таблицы в [db.ts](../../server/src/commercial/db.ts) | Хэшированные expiring fixture tokens, DB membership/active role, cross-org 404. Это test identity, не OIDC |
| Entitlements | [identity.ts](../../server/src/commercial/identity.ts) | Free/Pro/Business/Team как допустимые планы; active/expiry и hosted/BYOK flags проверяются в транзакции. План/роль/цена из body не принимаются |
| Spend | [ledger.ts](../../server/src/commercial/ledger.ts), [service.ts](../../server/src/commercial/service.ts) | Global/org/user счета; bigint micro-USD, reserving under row locks; idempotency hash; max-inflight; quote/output/time bounds |
| Recovery | [ledger.ts](../../server/src/commercial/ledger.ts) | Reserved без dispatch можно освободить; dispatched/unknown сохраняют reserve. Запрос, который мог быть оплачен, автоматически не повторяется |
| Usage | [providers.ts](../../server/src/commercial/providers.ts), existing [usageMetrics.ts](../../server/src/usageMetrics.ts) | Pricing version, actual provider cost и сервисные списания отдельно; usage нормализуется; prompts/tools не сохраняются в журнале расходов |
| BYOK / secrets | [secrets.ts](../../server/src/commercial/secrets.ts) | AES-256-GCM dev protector за интерфейсом SecretProtector; AAD org/key/provider/owner/version; private/org keys, rotate/revoke, hints only. Production требует отдельного managed KMS adapter |
| API безопасность | [app.ts](../../server/src/commercial/app.ts) | Loopback/Host/Origin, auth до JSON parser, 32 KiB body, строгие поля, 200 requests/min/user в процессе; monetary/concurrency limits в БД, не в Map |
| Эксплуатация | [migrate.ts](../../server/src/commercial/migrate.ts), [seed.ts](../../server/src/commercial/seed.ts), [compose](../../compose.commercial.yml), [CI](../../.github/workflows/commercial-ci.yml) | Checksummed transactional migration; rollback только пустой схемы; fixture credentials в ignored файле; health/readiness; sanitized JSON logs |

### Зависимости

Добавлены ровно два прямых пакета: runtime `pg@8.23.1` и dev `@types/pg@8.23.1`, оба MIT по метаданным npm. Цель — настоящий PostgreSQL с транзакциями/блокировками; in-memory имитация не доказывает concurrency correctness. ORM, Redis, billing/cloud SDK не добавлены. Версии закреплены в [server/package-lock.json](../../server/package-lock.json).

В том же серверном lockfile обновлены совместимые транзитивные исправления: `body-parser 1.20.8→1.20.9`, `node-forge 1.4.0→1.4.3`, `proxy-addr 2.0.7→2.0.8`, `qs 6.16.0→6.16.2`. Повторный `npm audit --prefix server` показывает **0 vulnerabilities** на дату проверки. Корневой dev lockfile не обновлялся: исторические root advisories C0 не считаются автоматически закрытыми.

PostgreSQL 17 для тестов закреплён digest в Compose/CI. Docker Desktop запущен для локальных тестов; контейнер `amai-c1-postgres` слушает только `127.0.0.1:55432`. В нём только синтетические development/test базы, не данные Excel.

## Доказательства

- [commercial.test.ts](../../server/src/commercial/commercial.test.ts): интеграционные тесты с настоящим PostgreSQL, HTTP server, несколькими pool/client и дочерним процессом с принудительным завершением после committed dispatch.
- 100 конкурентных новых запросов при бюджете на один запрос → один upstream attempt. 100 повторов одного idempotency key → один attempt и settlement; другое body → conflict.
- Изолированы global/org/user caps, BYOK service units, исходный billing period; role/license и credential version повторно проверяются перед dispatch.
- Проверены lost usage, abort/ошибка провайдера, pre-dispatch vault failure, overrun freeze, restart recovery, идемпотентная сверка и отказ double settlement.
- Cross-tenant/private keys, shared key permissions, ciphertext swap, key rotation/revocation, hints/redacted errors/logs, foreign Host/Origin, unknown body fields и custom provider URL проверяются негативными сценариями.
- Existing baseline: **659 тестов панели + 130 серверных** прошли. Основная и web-сборки успешны. Код `src/excel`, `src/agent` и текущий серверный entrypoint не изменены.
- Новый commercial suite: **19 тестов**. Полный итог прогона и CI следует проверять по последнему commit, а не только по числу test-файлов. Локальные журналы: `%TEMP%/amai-c1-tests.log`, `amai-c1-root-tests.log`, `amai-c1-server-tests.log`, `amai-c1-root-build.log`, `amai-c1-web-build.log`; они не коммитятся.
- Выполнен отдельный startup smoke: dev seed → отдельный HTTP server → authenticated fake chat → persisted settled result → graceful shutdown. Внешних AI-запросов и реальных списаний не было.

## Статус критериев C1

| Критерий | Состояние |
| --- | --- |
| AC-C1-01: separate entrypoint, local compatibility, no public paid mode | Проверено локально, CI предусмотрен |
| AC-C1-02: tenant isolation в реальной БД | Проверены keys, usage, memberships, composite references; cloud file/export endpoint в C1 отсутствует, поэтому не открыт без tenancy |
| AC-C1-03: конкурентный бюджет | Проверено на PostgreSQL: 100 конкурирующих обращений, один dispatch |
| AC-C1-04: dedupe/crash/disconnect | Проверено на fake provider и реальном process crash; реальный provider idempotency отдельно не заявлен |
| AC-C1-05: uncertain usage/pricing/period | Реализовано и проверено; reconciliation — внутренний trusted API, не endpoint с ценой от пользователя |
| AC-C1-06: vault/redaction/ops/dependencies | Dev vault, AAD/rotation/revoke, logs и миграции проверены; managed KMS и staging restore drill не выполнены. Server audit 0; root tooling triage остаётся |

## Ограничения и следующие интеграции

1. Нет настоящего login/OIDC — это C2. Включённый C1 намеренно не стартует в `NODE_ENV=production`. Не снимайте guard ради демонстрации.
2. Нет выбранного managed KMS и инфраструктурной identity; dev encryption key задаёт оператор. Потеря его значения делает тестовые ciphertext недоступными. Нельзя переносить dev master key в публичное окружение.
3. Ни один реальный provider не включён в commercial registry; local edition сохраняет прежние 9 провайдеров. Подключение real adapter требует bounded pricing/output/retry policy, ключа платформы/tenant secret и согласованного малого бюджета; новые реальные расходы в этой задаче не разрешались.
4. Нет public org creation/invites/seats/billing, тарифных цен или автоматической выдачи credits. Development fixture — только явно вызванный seed. Базовые entitlement flags не заменяют продуктовую SKU policy C2.
5. 200 requests/min/user — локальный process limiter, не распределённый abuse-control. Межпроцессные money/concurrency caps и dedupe уже в PostgreSQL. Полная нагрузочная приёмка публичного gateway — C5.
6. SSE C1 — закрытый тестовый протокол request/delta/done; commercial transport в Excel ещё не подключён. Журнал Excel-операций и rollback — C3, здесь не внедрялись.
7. DB recovery из резервной копии, метрики 30-day SLO и выбранный hosting не проверены. Запуск Docker/test DB не является production deployment.

Запуск и действия оператора: [C1_RUNBOOK.md](C1_RUNBOOK.md). Риски C0 остаются открытыми вне доказанного test scope. После review этой поставки можно подключать C2 identity и выбирать staging/KMS, сохраняя default zero-spend.
