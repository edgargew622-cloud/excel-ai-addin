# am.AI — план C1–C6

Исходный план: 7 ноября 2026. **8 ноября пользователь разрешил начало C1 по архитектуре C0**; кодовая база реализации — опубликованная 1.0.48. Основание: [C0_AUDIT](C0_AUDIT.md), [SAAS_ARCHITECTURE](SAAS_ARCHITECTURE.md), [RISK_REGISTER](RISK_REGISTER.md). Статус уже созданных компонентов и оставшихся gates — в [C1_PROGRESS.md](C1_PROGRESS.md); остальные файлы ниже по-прежнему планируемые.

## Правила выполнения и оценка

- C0 — только пять документов. Перед C1 — утверждение архитектуры и использования `production`, поскольку `main` отсутствует.
- Отдельные PR с наблюдаемыми критериями; существующие Excel-инструменты и модели сохраняются. При будущей необходимости изменения shared transport/observer сначала тесты совместимости, затем ограниченный PR.
- Новая зависимость требует описания цели, альтернативы, license/security review и locked версии. Никаких live payment keys в Git.
- Ни этап, ни его приёмка сами по себе не разрешают production deployment или merge: требуется отдельное утверждение. Staging с синтетическими данными; реальный расход только с согласованным бюджетом.
- Оценки — **человеко-дни** опытной команды, 6–8 продуктивных часов/день; не календарное обещание. Включены код, тесты, review/docs; не включены ожидание merchant/store approval, поиск design partners и юридические работы. Один человек выполняет этапы последовательно; номинальный итог **88–139 человеко-дней**, резерв 25–35% даёт **110–188**.
- Порядок milestones: C1 → C2 → C3 → C4 → C5 → C6. Часть исследований может пересекаться только после утверждения владельца.

## C1 — коммерческий backend (18–28 дней)

**Зависимости:** утверждённый C0, тестовый hosting/DB/secret store, zero-spend default и ограниченный pilot budget, выбранная схема tenant data. Продуктовые числа не нужны для выключенного hosted режима; нужны до его включения.

**Работы:** отдельный Express entrypoint, config/env separation local/cloud, API versioning, repositories/migrations, scoped RequestContext, entitlements skeleton, budget accounts + atomic reservation, usage attempts/ledger/reconciliation, tenant vault, provider credential injection, gateway health и structured redacted logs. C1 использует только проверяемую тестовую identity в non-production; сетевой paid endpoint остаётся закрыт до C2.

**План файлов:** новые `server/src/commercial/server.ts`, `config.ts`, `requestContext.ts`, `db/{migrations,repositories}`, `entitlements/{policy,plans}.ts`, `usage/{ledger,reservations,reconcile,pricing}.ts`, `secrets/credentialStore.ts`, `providers/adapter.ts`. Переиспользовать [protocol](../../server/src/protocol.ts), [responsesApi](../../server/src/responsesApi.ts), [openaiRoute](../../server/src/openaiRoute.ts), [usageMetrics](../../server/src/usageMetrics.ts), [providers](../../server/src/providers.ts). Не менять Excel executor. Runtime integration adapter отделяет глобальные local callbacks от tenant-scoped credentials.

**Тесты:** новая commercial unit/integration suite на настоящем PostgreSQL; два org, foreign ID/ref/list/export denial; 100 конкурентных reservations на малом budget; dedupe request/body hash; crash до/после dispatch; неизвестный usage, provider/SDK retry, отмена stream; vault rotate/delete/redaction; existing `npm run check` и web build.

**Приёмка:** AC-C1-01…06 из [ACCEPTANCE_CRITERIA](ACCEPTANCE_CRITERIA.md); без valid reservation нет upstream; БД сохраняет pending/settled после рестарта; global/user/org caps не обходятся параллелизмом; отключённый commercial mode не меняет local startup. Записать performance baseline, rollback миграции и runbook. **Выход:** staging backend с fake auth, без публичной платной эксплуатации.

## C2 — авторизация и подписки (12–18 дней)

**Зависимости:** C1; выбранный identity provider, registered redirect origins; утверждённые роли/Free–Pro–Business–Team entitlements; доступ к тестовому Windows/Mac/Online. Платёжный merchant пока не нужен.

**Работы:** identity verification, PKCE/dialog login, безопасная сессия и logout/revoke, personal org, invitations/roles/seats, membership switching, server license checks, fake BillingProvider/manual test entitlements, коммерческий transport taskpane за feature flag. Offline local режим остаётся явно отдельным.

**План файлов:** новые `server/src/commercial/auth/{identity,sessions,authorization}.ts`, `organizations/{routes,service}.ts`, `billing/{provider,fakeProvider}.ts`, `src/taskpane/api/commercialClient.ts`, `src/taskpane/auth/*`; ограниченная интеграция [api/client.ts](../../src/taskpane/api/client.ts), [Taskpane.tsx](../../src/taskpane/Taskpane.tsx), hosted manifest/redirect config. Pure Excel modules не изменяются.

**Тесты:** issuer/audience/expiry/signature, неверный state/nonce, refresh/revoke, CSRF/Origin в выбранной session схеме, IDOR между org, last-owner, duplicate invite/seat race, expired entitlement и смена org во время SSE. Login/logout/reload на трёх настоящих host. Прямой API без лицензии/с неверной ролью не вызывает provider; UI обходится в негативных тестах намеренно.

**Приёмка:** AC-C2-01…05; нет identity заглушек в production config; нет refresh token в localStorage/URL; revoked membership теряет новые paid requests; недоступность auth не запрещает локальную отмену. **Выход:** закрытая авторизованная staging beta с fake subscription states.

## C3 — надёжность Excel-операций (18–28 дней)

**Зависимости:** C2 для tenant журналов, доступ к реальным host; заранее утверждённая recovery/retention policy. Исследование host export capabilities можно начать после C0, реализацию — после утверждения.

**Работы:** durable local operation intent, bounded sync queue, append-only server journal, operation IDs и state transitions, crash reconciliation, безопасная компенсирующая отмена, backup/restore UX для irreversible операций, explicit capabilities. Существующий undo остаётся, его ограничения показываются. Instrumentation через adapters/observers; изменение алгоритмов Excel — только отдельным согласованным исправлением выявленного дефекта.

**План файлов:** новые `src/excel/operationJournal.ts`, `src/taskpane/operationRecovery.ts`, `server/src/commercial/operations/{routes,store}.ts`, `scripts/commercial-excel-smoke.mjs`, `docs/commercial/platform-matrix.md`; интеграционные точки [loop](../../src/agent/loop.ts), [plans](../../src/excel/plans.ts), [undo](../../src/excel/undo.ts), [backup](../../src/excel/workbookBackup.ts), [workbook events](../../src/excel/workbookEvents.ts). Добавление journal callbacks не должно менять write/verify semantics.

**Тесты:** crash до write/после write/до verify, reload, double submit, cancel при partial group, change workbook, newer manual edit, coauthor changes и отсутствующий event monitor; журнал дубликатов/out-of-order; disk/quota/network failure; restore xlsx/xlsm с несохранёнными изменениями и внешними ссылками; отдельно отказ host export API.

**Приёмка:** AC-C3-01…06; каждый baseline critical scenario на Windows/Mac/Online имеет evidence SHA+host build; unknown никогда автоматически не повторяет write; проверенное восстановление либо явная блокировка неподдержанной destructive операции; денежный/лицензионный отказ не ломает already-started write и rollback. **Выход:** recovery-ready beta, без обещания универсального ACID/полного undo.

## C4 — платежи и биллинг (10–16 дней)

**Зависимости:** C1 ledger, C2 identity/entitlements, утверждённые SKU/цены/валюта/налоги/merchant geography. Выбранный BillingProvider, только sandbox credentials.

**Работы:** конкретный adapter за интерфейсом, checkout/portal, signed webhook raw body, durable inbox/outbox, subscriptions projection, reconciliation, plan/seat lifecycle, refund/past_due policy. Customer receipts и финансовый учёт отделены от provider usage ledger.

**План файлов:** новые `server/src/commercial/billing/{routes,webhooks,inbox,reconcile,adapters/*}.ts`, migrations billing tables, `src/taskpane/billing/*`, `docs/commercial/billing-runbook.md`. Из C2 переиспользуются `provider.ts` и fake provider; AI adapters не меняются.

**Тесты:** invalid signature, stale timestamp, duplicate/out-of-order/delayed webhook, provider downtime, replay после restart, duplicate checkout, unpaid redirect, cross-org portal, возврат и отмена, prorations/seats, renewal/period rollover и уже зарезервированный usage. Sandbox fixtures не содержат реальных карт/ключей.

**Приёмка:** AC-C4-01…05; одна оплата/событие применяется один раз, redirect не даёт лицензию, reconciliation восстанавливает пропущенное состояние, сумма charge отличается от estimate явно. **Выход:** sandbox billing готов; live activation только после отдельного разрешения.

## C5 — публичная бета и публикация (15–25 дней)

**Зависимости:** C1–C4 приняты; критичные launch risks закрыты или явно согласован ограниченный scope. Получены согласия на production и отдельно на live payments/store submission; подтверждены merchant и privacy terms. Конкретный номер версии утверждается отдельно, C0 не повышает `1.0.46`.

**Работы:** release candidate, platform matrix, dependency/license/SBOM audit, key/secret scan, rate/abuse/load tests, accessibility/onboarding, privacy/export/delete, staged rollout и kill switch, restore drill/incident runbooks, support и beta feedback. Настройка публикации без автоматического переноса нынешних production/tag triggers на commercial target.

**План файлов:** новые `.github/workflows/commercial-ci.yml`, `commercial-release.yml`, hosted manifest/config, `docs/commercial/{release-checklist,incident-runbook,platform-matrix}.md`; интеграция с [windows workflow](../../.github/workflows/windows-bundle.yml), [web workflow](../../.github/workflows/web-panel.yml), [packaging](../../packaging/ExcelAI.iss), [bundle smoke](../../scripts/bundle-smoke.ps1), [INSTALL](../INSTALL.md). Существующая local поддержка сохраняется.

**Тесты:** clean install/update/uninstall Windows, reauth, true Online/Mac runtime, потеря сети/лимит подписки, данные двух tenants, 2× согласованной beta concurrency, SQL backup restore и billing/usage reconciliation. 14 дней закрытого pilot наблюдения — предварительный gate, календарно дополнительно к трудозатратам.

**Приёмка:** AC-C5-01…06; фактические SLO/latency замерены, известные исключения опубликованы; нет нерешённых подтверждённых critical/high на достижимом commercial path; explicit release approval. **Выход:** ограниченная публичная beta; само наличие checklist ничего не публикует.

## C6 — аналитика и масштабирование (15–24 дня)

**Зависимости:** C5 и реальные согласованные privacy-safe telemetry; минимально достаточные данные о нагрузке/экономике. Без измерений не добавлять Kubernetes/Redis/очереди.

**Работы:** activation/retention, conversion/ARPA, hosted-vs-BYOK usage, provider spend/gross margin, unknown usage и support burden; PII-minimization и retention. Оптимизация prompt/cache/router при сохранении качества; горизонтальные gateway replicas при необходимости, единый ledger/concurrency, фоновые jobs с leases. Сценарии масштабирования выбираются по bottleneck.

**План файлов:** `server/src/commercial/analytics/*`, `observability/*`, load fixtures, dashboards/config, `docs/commercial/{unit-economics,capacity-plan}.md`; переиспользовать [usageMetrics](../../server/src/usageMetrics.ts), [promptCache](../../server/src/promptCache.ts), добавить exporter только при выборе backend.

**Тесты:** воспроизводимый load/restart/failover, quota при нескольких replicas, lease fencing, reconciliation duplicate delivery, metrics redaction, pricing/catalog version changes и quality regression corpus Excel-задач.

**Приёмка:** AC-C6-01…04; dashboard сходится с ledger, scale не меняет correctness, стоимость одного успешного сценария измерена, данные для $1 млн valuation не подменяются количеством функций. **Выход:** проверенный capacity plan и экономические факты, без обещания оценки компании.

## Первая поставка после утверждения

Начинать C1 с одного закрытого vertical slice: **test identity → org → entitlement → atomic budget reservation → fake provider stream → durable usage settlement**. Затем tenant-isolated secret lookup и один реальный adapter в staging с малым бюджетом. Прежде чем подключать платежи и новые Excel-возможности, доказать невозможность cross-tenant доступа и перерасхода при параллельных запросах.
