# C2 — закрытый пилот через AI Gateway

Дата: **9 ноября 2026, Europe/Moscow**. Пользователь разрешил C2 и выбрал **Microsoft Entra ID**. Ветка `codex/commercial-c2`, отдельный worktree; исходная рабочая папка и незакоммиченные изменения пользователя сохранены.

## База и текущий статус

Проверенные удалённые refs перед реализацией:

| Ref | Commit |
| --- | --- |
| `production`, 1.0.46 | `930e8411ac2cb1da1af523273cf1a1da74e77e1d` |
| `v1.0.48^{}` — commit тега, версия 1.0.48 | `23830108707dbb1856537b1e22798cb25b164631` |
| `codex/commercial-c0` | `db262cfc17974fded53771d26dbd565943cc61f7` |
| `codex/commercial-c1` — **базовый SHA C2** | `355f2cfa65e15fc6c3d38887db0c65c02d97a861` |

У аннотированного тега `v1.0.48` собственный Git object SHA `7f5e5090c2a6d5547b36fe817a17a13aba5c6457`; его нельзя путать с commit выше. C1 CI на базе: [успешный run](https://github.com/edgargew622-cloud/excel-ai-addin/actions/runs/42420292381). Итоговый SHA поставки: `git rev-parse HEAD` в ветке C2; он также указан в итоговом отчёте и Actions run этой ветки. Версия приложения оставлена **1.0.48**, новый релиз не создаётся.

**Статус: реализация для проверки; C2 ещё не принят.** Подписанные тестовые OIDC-токены и имитатор API не доказывают вход в действующий Entra tenant. Настоящий вход и сквозной запуск в Excel Windows ожидают регистрации приложения/локального секрета и проверки на host. Платная модель не включена; реальная себестоимость не измерена.

## Изменения и доказательства

| Область | Код и проверка |
| --- | --- |
| Исправление адреса PostgreSQL | [core.ts](../../server/src/commercial/core.ts): проверка URI до **любого** `Pool`, запрет query/hash и удалённого hostname. `host`, encoded/duplicate host, hostaddr, service, port, ssl options проверяются без сетевых соединений в [c2.test.ts](../../server/src/commercial/c2.test.ts). Не только `readConfig`, но и вспомогательные CLI/test paths проходят общий guard |
| Реальная identity integration | [auth.ts](../../server/src/commercial/auth.ts), [c2Server.ts](../../server/src/commercial/c2Server.ts): Entra single tenant, code + PKCE S256, state/nonce, RS256/JWKS, issuer/audience/exp/iat/tid/oid, multi-audience azp. Это рабочая интеграция, ещё не проверенная с действующей регистрацией |
| Сессия | Одноразовый callback claim в БД, parent-bound redemption ticket, hashed expiring session. Токены IdP остаются на сервере; refresh token не запрашивается. В панели app bearer только в памяти; после reload требуется новый Office Dialog/Entra SSO. Logout отзывает session; новый запрос проверяет active user/membership/license |
| Приглашение | [provisionPilot.ts](../../server/src/commercial/provisionPilot.ts): оператор связывает существующий **синтетический** user с verified Entra issuer + Object ID. Сам вход не создаёт org, подписку или credits. Это операторский allowlist, не UI invitations/seats |
| Gateway transport | [commercial.ts](../../src/taskpane/api/commercial.ts), [client.ts](../../src/taskpane/api/client.ts), [Taskpane.tsx](../../src/taskpane/Taskpane.tsx): build flag и выключенный checkbox; same-origin `/api/v1`; общий existing SSE parser; `tool_calls` и tool results; org selector; состояние и тестовый бюджет |
| Исполнитель Excel | [loop.ts](../../src/agent/loop.ts): инъекция per-task chat, отключение network retries только для commercial task, исключение local-service tools из предложения, уведомление об остановке. Existing plan/confirm/read scope/read-back/partial-write/undo paths сохранены; файлы `src/excel/*` не изменены |
| Каталог | [toolCatalog.json](../../server/src/commercial/toolCatalog.json) генерируется из существующего registry [скриптом](../../scripts/commercial-tool-catalog.mjs). Сервер принимает только имена разрешённых tools, схемы и descriptions берёт из своего catalog. Службы локальных файлов/памяти/интернета не превращаются в облачные endpoints. Drift проверяет [commercial.test.ts](../../src/taskpane/api/commercial.test.ts) |
| Учёт задания | [tasks.ts](../../server/src/commercial/tasks.ts), [ledger.ts](../../server/src/commercial/ledger.ts), [schemaV2.ts](../../server/src/commercial/schemaV2.ts): task ID, org/user ownership, max steps/time/spend, sequential dispatch, durable spent/reserved. Существующие global/org/user reservations и отдельные hosted/BYOK meters сохранены. Попытки связаны с заданием через usage request/FK |
| Обрывы | Потерянный ответ не вызывает автоматический платный retry или Excel replay. Повтор request ID возвращает только metadata; parser отвергает её как исполняемый ответ. Последний org/task pointer сохраняется в sessionStorage **без** bearer, prompts, cell values. После reload — вход и кнопка состояния, никакого автоматического продолжения |
| Один реальный adapter | [deepseekPilot.ts](../../server/src/commercial/deepseekPilot.ts): подготовлен кандидат DeepSeek Chat Completions, фиксированный URL, finite input/output bound, server price version/expiry, bigint micro-USD, single HTTP attempt, missing usage → unknown. HostedSecretSource и tenant credential разделены. В runtime registry C2 **только fake**. Выбор реальной модели и цен требует отдельного утверждения |
| Миграция/безопасность | V1 не переписана, добавлена checksummed V2. Empty rollback теперь также отказывается удалять пользователей/auth flows. C2 не принимает C1 fixture bearer, сохраняет localhost DB/bind и production guard. Структурированные технические логи — только метаданные |

## Проверки

Команды и воспроизводимый Windows-сценарий — в [C2_RUNBOOK.md](C2_RUNBOOK.md). Запускайте destructive fixture suites **последовательно**, либо `node --test --test-concurrency=1 ...`; они используют одну изолированную тестовую БД и очищают схему. Не запускайте два suite одновременно в одной БД.

- Root baseline + новые transport tests: **665** тестов (659 прежних + 6 новых).
- Existing server suite: **130** тестов.
- Commercial C1 regression: **19** тестов.
- Commercial C2: **11** интеграционных тестов с настоящим PostgreSQL, криптографически подписанными JWT и имитатором IdP/model HTTP.
- Основная сборка, web-сборка и отдельная сборка с `VITE_COMMERCIAL_C2=1` проверяются раздельно.
- [CI](../../.github/workflows/commercial-ci.yml) запускает обе PostgreSQL suites последовательно, существующие server/root tests и три варианта сборки. Статус конкретного SHA нужно проверять в GitHub Actions; наличие workflow не является доказательством успешного run.
- Тесты охватывают state/nonce/PKCE/подпись/issuer/audience/expiry/tenant, parent origin/proof/replay, session expiry/revoke, foreign tenant, role/subscription, task budget/time/step/mode, BYOK service units, 100 duplicate requests, restart/uncertain funds, tool protocol, lost SSE, no automatic retry, secret-log sentinels. C1 отдельно сохраняет real child-process crash test и 100 competing distinct requests.

### Зависимости

Добавлен один runtime пакет **`jose@6.2.16` (MIT)**, версия закреплена: проверка JWT/JWKS вместо самописной криптографии. Express/pg/существующий протокол переиспользованы; auth SDK в клиент не добавлен. `npm --prefix server audit` на момент проверки — **0 vulnerabilities**. Это не закрывает прежние root tooling advisories и не является полным security audit. Новые cloud/billing зависимости не добавлены.

## Открытые критерии и точные следующие действия

1. Владелец регистрирует single-tenant **Web** application Entra и redirect URI, задаёт tenant/client IDs и server-only secret локально. Подробные шаги и allowlist — [runbook §2](C2_RUNBOOK.md#2-microsoft-entra-id).
2. Проверить вход/выход/повторный вход после reload на конкретном Excel Windows, записать version/build/SHA и результат [сценария](C2_RUNBOOK.md#5-проверка-в-excel-windows). До этого основной gate C2 **pending**.
3. Тарифная матрица, места/invites/last owner, коммерческие цены и квоты не утверждены. В коде только finite synthetic pilot policy, не продуктовые тарифы.
4. Для выхода за пределы localhost отдельно выбрать hosting, TLS origin и managed secret store/IAM; заменить dev protector, проверить восстановление БД и доступность старых key versions. Публичный bind/удалённая БД/production сейчас намеренно закрыты.
5. Для реального AI отдельно утвердить провайдера/модель, срок и snapshot цен, global/org/user/task budgets и разрешение на конечный тестовый расход. Затем подключить HostedSecretSource, зарегистрировать adapter и выполнить ограниченный smoke. До этого live AI и payments выключены.
6. Состояние `completed` означает завершение агентского протокола, не серверное доказательство содержимого Excel. Клиентский executor/read-back и реальный host smoke остаются обязательны. Журнал crash-safe Excel intents/undo — C3.

Ни merge, ни релиз, ни production deployment не выполняются в C2. [Реестр рисков](RISK_REGISTER.md) остаётся действующим; не закрывайте риски по одному fake-сценарию.
