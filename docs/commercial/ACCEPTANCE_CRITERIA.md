# am.AI — критерии приёмки C0–C6

Дата: 7 ноября 2026. Основание: [аудит](C0_AUDIT.md), [архитектура — проект](SAAS_ARCHITECTURE.md), [roadmap](ROADMAP.md), [риски](RISK_REGISTER.md). Статусы здесь не заменяют результаты тестов. C1–C6 **не выполнены** и не утверждены.

## Формат доказательств

Для каждого gate: SHA, версия/feature flags, среда и host build, команда/сценарий, expected/actual, время, обезличенный test artifact и ответственное лицо. Для Excel дополнительно fixture checksum, before/after/формулы и unaffected ranges; для затрат — fake provider trace, reservation и ledger IDs. Нельзя публиковать реальные keys, full workbook paths, данные клиентов и приватные prompts в CI artifacts. Ссылка на test-файл доказывает наличие проверки, успешный run — отдельное evidence.

Доказательства baseline: [640+130 tests, builds/manifest](C0_AUDIT.md). Критические baseline invariants: [toolSchemas.test](../../src/excel/toolSchemas.test.ts), [guarantees.test](../../src/agent/guarantees.test.ts), [undo.test](../../src/excel/undo.test.ts), [readScope.test](../../src/agent/readScope.test.ts), [localOnly.test](../../server/src/localOnly.test.ts), [usageMetrics.test](../../server/src/usageMetrics.test.ts).

## C0 — документация и граница утверждения

| ID | Критерий | Статус C0 |
| --- | --- | --- |
| AC-C0-01 | Зафиксированы actual default ref/SHA/версии; отсутствие `main` раскрыто | Выполнено: `production`, `930e8411ac2cb1da1af523273cf1a1da74e77e1d`, 1.0.46 |
| AC-C0-02 | Все 5 запрошенных документов существуют, внутренние ссылки разрешаются; facts отделены от proposals | Выполнено в ветке `codex/commercial-c0` |
| AC-C0-03 | Нет diff рабочих функций/dependencies/manifests/workflows; dirty stage10 не включён; нет deployment/merge/payment keys | Проверяется Git diff относительно baseline; разрешённый список — только 5 `docs/commercial/*.md` |
| AC-C0-04 | Повторены tests/build/web/manifest, результаты не смешаны с dirty workspace | Выполнено, 770 тестов passed, команды в audit |
| AC-C0-05 | Владелец утверждает архитектуру и замену отсутствующей `main` на `production` | **Ожидает утверждения; C1 не начинать** |

## C1 — backend

| ID | Проверяемый результат | Evidence / gate |
| --- | --- | --- |
| AC-C1-01 | Отдельный commercial entrypoint; local mode совместим; production paid AI выключен без identity integration | Existing `npm run check`, web build и новые config tests; diff Excel tools отсутствует |
| AC-C1-02 | Два tenants не видят чужие credentials, rows, lists, operation/usage refs; foreign composite refs отвергаются | DB integration tests с реальными scoped repositories, не только mocks; R01/R03 |
| AC-C1-03 | 100 конкурентных запросов при budget на 1 запрос создают максимум 1 допустимый upstream reservation/dispatch; org/user/global budgets и cap периода соблюдаются | Transactions/concurrency тест на PostgreSQL; `spent+reserved≤cap`, не одно только HTTP 429; R02 |
| AC-C1-04 | Повтор ID+body не удваивает charge/dispatch; другой body с тем же ID → 409; restart после dispatch не делает автоматический replay | Fault injection before/after dispatch, timeout, SSE disconnect, route/SDK retries; R02/R12 |
| AC-C1-05 | Missing usage остаётся uncertain и резервом, не нулём; reconciliation идемпотентен; pricing version/currency закреплены | Ledger trace и provider stubs; restart/period rollover tests, zero-spend default |
| AC-C1-06 | Tenant key context, rotate/revoke и redaction работают; DB rollback/runbook есть; dependency advisories triaged | Canary secrets отсутствуют в logs/bundle/responses, vault isolation tests; R03/R07/R11/R14 |

## C2 — identity и лицензии

| ID | Проверяемый результат | Evidence / gate |
| --- | --- | --- |
| AC-C2-01 | Invalid signature/issuer/audience/expired auth и неверный dialog state/nonce отвергаются; production не принимает fake identity | Auth negative integration suite; три host login traces |
| AC-C2-02 | Free/Pro/Business/Team и роли применяются server-side; изменение plan/org/user в body или UI не даёт доступ | Direct API negative tests, истёкшая/отозванная лицензия, R01 |
| AC-C2-03 | Invite/seat race не превышает число мест; viewer не тратит org budget; последний owner защищён | Два org, concurrent invite accept, last-owner tests |
| AC-C2-04 | Logout/revoke/membership removal запрещают новый AI request; auth outage не запрещает local undo/recovery | Session lifecycle tests; refresh credentials не в localStorage/URL, выбранная cookie схема проверена на CSRF |
| AC-C2-05 | Existing local/direct BYOK сохранён отдельно; commercial mode идёт через gateway и не выдаёт key клиенту | Transport regression suite и disclosure в UI; нет обещания оплаты за неуправляемый offline-код |

## C3 — данные Excel и recovery

| ID | Проверяемый результат | Evidence / gate |
| --- | --- | --- |
| AC-C3-01 | Любой mutating tool имеет plan/confirm/analysis-only guard; изменения runtime не обходят baseline guarantees | Полный registry test + реальные write/decline/read-back smoke; R04 |
| AC-C3-02 | Crash/reload до/после write ведёт к durable intent и unknown/reconcile; автоматический replay unknown запрещён | Fault matrix `prepare/approve/write/read-back/journal-sync`; R04 |
| AC-C3-03 | Partial group не показана полным успехом; cancel завершает начатую безопасную часть и не запускает следующую | [guarantees](../../src/agent/guarantees.test.ts) + real Excel before/after; нет повреждений unaffected ranges |
| AC-C3-04 | Undo не затирает новую ручную/coauthor правку; structural invalidation и недоступный monitor видны пользователю | Real conflict fixtures на hosts, rollback_conflict; R05/R06 |
| AC-C3-05 | Для irreversible operations есть проверенный recovery path либо явная блокировка unsupported commercial действия; license expiry не блокирует recovery | Restore drill xlsx/xlsm; host export failure case; описаны ограничения unsaved edits/retention, R05/R09 |
| AC-C3-06 | Journal не содержит cell values/full path; event IDs/sequence dedupe; offline bounded queue не теряет intent молча | Privacy sentinels + restart/out-of-order tests; Windows/Mac/Online platform matrix, R08/R10 |

Минимальная host matrix: чтение/формула/обычная запись/format, сортировка/filter, insert/delete rows/columns, copy/delete sheet, table, pivot/chart, import, backup capability, cancel, partial group, stale state, undo, login/logout и budget rejection. Запись host+build+SHA обязательна. Unsupported функции явно помечаются и ограничиваются capability gate; WebKit smoke и macOS install script не засчитываются как настоящий Excel E2E. Существующие доказательства и ограничения — [audit §4–5](C0_AUDIT.md).

## C4 — billing sandbox

| ID | Проверяемый результат | Evidence / gate |
| --- | --- | --- |
| AC-C4-01 | Billing domain не зависит от конкретного merchant SDK; fake и sandbox adapters проходят один contract suite | Interfaces и tests, нет реальных payment keys |
| AC-C4-02 | Невалидная/просроченная подпись отклоняется по raw body; duplicate/out-of-order events корректны | Webhook replay/concurrency suite; R13 |
| AC-C4-03 | Browser success redirect/поддельный metadata org не выдают license; только доверенная server projection | Direct API/redirect negative tests |
| AC-C4-04 | Renewal/cancel/refund/past_due/seat changes и billing period корректны; reserve из предыдущего периода не теряется | Sandbox lifecycle + ledger reconciliation tests |
| AC-C4-05 | Provider outage/потерянный webhook/restart восстанавливаются reconciliation без double charge | Durable inbox/outbox trace и runbook; live charging остаётся выключенным |

## C5 — release/public beta

| ID | Проверяемый результат | Evidence / gate |
| --- | --- | --- |
| AC-C5-01 | Подтверждённых достижимых Critical/High нет; все исключения имеют ограниченный scope, evidence и явное принятие владельцем | Новый audit/SBOM/secret scan на release SHA, triage R07; никаких выводов «0 CVE» из одного успешного build |
| AC-C5-02 | Actual Windows/Mac/Online matrix пройдена; clean install/update/uninstall совместимы с поддержанными конфигурациями | Записанные host runs и Windows packaging smoke; непроверенные платформы не рекламируются как проверенные |
| AC-C5-03 | Finite plan/global budgets включены, rate/concurrency/abuse limits проверены при 2× заявленной beta нагрузки | Load profile с числом клиентов/RPS/длиной stream; нельзя проходить gate без конкретного профиля |
| AC-C5-04 | Privacy/export/delete/retention, merchant feasibility, license notices и support готовность подтверждены ответственными | Ссылки на reviewed policies и tests; не самообъявленное compliance |
| AC-C5-05 | Metadata DB restore проверен: предложенные RPO≤24ч/RTO≤4ч; kill switch и incident runbook работают | Staging restore timestamps, usage/billing сверка; recovery текущей Excel-книги оценивается отдельно |
| AC-C5-06 | Предлагаемые SLO 99,5% gateway/30 дней, p95 auth+quota≤500мс и alerts≤5мин приняты; pilot evidence собраны; релиз одобрен явно | Закрытый pilot ≥14 дней как предварительный gate; итоговая версия и scope утверждены; без разрешения нет deployment/merge/store submission |

Полный 30-дневный SLO нельзя объявить достигнутым по 14-дневному pilot: на старте beta это утверждённая цель, затем rolling measurement. Не зависеть от provider latency для измерения собственного overhead.

## C6 — измеримость и рост

| ID | Проверяемый результат | Evidence / gate |
| --- | --- | --- |
| AC-C6-01 | Активность/конверсия/retention и hosted/BYOK economics определены без записи содержимого книг | Reviewed event schema, retention, consent; метрики не смешивают provider cost и customer revenue |
| AC-C6-02 | Usage/revenue dashboard сверяется с ledger и billing; unknown/mismatch имеют отдельные счётчики | Reconciliation report, pipeline duplicate/restart tests |
| AC-C6-03 | Несколько replicas/failover не меняют budget/idempotency/tenant guarantees | Multi-instance concurrency/load/fencing suite; масштабирование только после baseline bottleneck |
| AC-C6-04 | Capacity plan и стоимость успешной Excel-задачи измерены; качество не падает при оптимизации router/cache | Versioned quality corpus и unit-economics report; оценка компании не считается техническим acceptance gate |

## Решение о начале C1

До начала реализации владелец утверждает [архитектурные решения](SAAS_ARCHITECTURE.md) и actual base `production`. C0-документы готовы к этому review; C1–C6 остаются планом. Первый PR после утверждения: закрытый vertical slice с fake provider и доказанными tenant isolation / atomic spend reservation.
