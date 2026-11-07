# am.AI — C0: технический аудит

Дата проверки: **7 ноября 2026, Europe/Moscow**. Статус: аудит выполнен; архитектура предложена, не утверждена. Изменений runtime, зависимостей, платежей и production нет.

## 1. Проверяемый снимок и ограничения

| Поле | Зафиксированное значение |
| --- | --- |
| Репозиторий | `edgargew622-cloud/excel-ai-addin` |
| Запрошенная ветка | `main` **отсутствует** в `git ls-remote --heads origin` |
| Фактическая default branch | `production`, подтверждено GitHub Repository API и удалёнными refs |
| Проверенный SHA | `930e8411ac2cb1da1af523273cf1a1da74e77e1d` |
| Commit date | `2026-10-29T14:18:10+03:00` |
| Версия панели / сервера | `1.0.46` / `1.0.46` |
| Версия manifest | `1.0.46.0` |
| Ветка документов | `codex/commercial-c0`, создана от указанного `origin/production` |
| Среда повторной проверки | Windows, Node `v24.16.0`, npm `11.13.0` |

Ссылка на [неизменяемый исходный снимок](https://github.com/edgargew622-cloud/excel-ai-addin/tree/930e8411ac2cb1da1af523273cf1a1da74e77e1d). Версии: [package.json](../../package.json), [server/package.json](../../server/package.json), [manifest.xml](../../manifest.xml). Относительные ссылки ниже относятся к этому дереву до добавления C0-документов.

Первоначальный рабочий каталог находится на `stage10`, SHA `030e47bc37883896fd17269028d4bdc5be4b8131`, с 21 изменённым tracked-файлом. Эти изменения **не входят** в аудит и не переносились. Создан отдельный Git worktree. Предыдущая оценка локального каталога, включая полуручной режим и 782 теста, не является оценкой `production`: в проверяемом снимке нет `approvalMode.ts`, все изменяющие операции требуют подтверждения. При утверждении C0 нужно подтвердить использование `production` вместо отсутствующей `main`.

Метод: чтение кода, тестов и workflow, воспроизводимые локальные проверки. Не проводились атаки на рабочие книги, запросы к платным LLM, установка приложения, проверка реального Excel, полный поиск секретов в истории Git, юридическая экспертиза или независимый pentest. Процент покрытия строк/ветвей не измерялся: текущие npm scripts не включают сбор coverage. Отсутствие функции ниже означает отсутствие в проверенном серверном/клиентском коде, а не во внешней инфраструктуре, доступ к которой не предоставлялся.

## 2. Воспроизводимые проверки

Зависимости установлены по существующим lockfiles: `npm ci --ignore-scripts --no-audit --no-fund` и аналогичная команда с `--prefix server`. Новых зависимостей нет; lockfiles не изменены. Lifecycle scripts при установке намеренно не выполнялись; npm scripts проверок запускались явно.

| Команда | Результат нового прогона на проверенном SHA |
| --- | --- |
| `npm run check` | Exit 0; **640 тестов панели + 130 серверных**, 0 failed; сервер компилируется внутри server test; TypeScript панели и Vite build успешны |
| `npm run build:web` | Exit 0; созданы web bundle, manifest и страница установки Mac |
| `npm run validate` | Exit 0; manifest valid; это проверка манифеста, не runtime совместимости |
| `npm audit --json` | 29 уязвимых пакетов: 4 critical, 18 high, 7 moderate; включая dev/transitive |
| `npm audit --prefix server --json` | 4: 1 critical, 2 high, 1 moderate |

Счётчики root/server audit пересекаются, их нельзя складывать. Audit — результат advisory-сервиса на дату проверки, не доказательство достижимости уязвимости. Зафиксированный серверный lock: `express 4.22.3`, `body-parser 1.20.8`, `proxy-addr 2.0.7`, `node-forge 1.4.0`, `qs 6.16.0`. Audit сообщает наличие исправлений; обновление major-пакетов требует отдельной проверки совместимости. Источники: [root lock](../../package-lock.json), [server lock](../../server/package-lock.json). Advisory IDs и условия: [реестр рисков R07](RISK_REGISTER.md).

Сводка stdout: `tests 640 / pass 640 / fail 0`, `tests 130 / pass 130 / fail 0`, `CHECK_EXIT=0`, `WEB_EXIT=0`, `MANIFEST_EXIT=0`. Полные локальные журналы: `%TEMP%/amai-c0-check.log`, `amai-c0-web.log`, `amai-c0-manifest.log`, `amai-c0-audit-root.json`, `amai-c0-audit-server.json`. Журналы не добавлены в Git. Воспроизводимые команды определены в [npm scripts](../../package.json) и [server scripts](../../server/package.json).

Сборщик предупреждает о неработающем разделении ряда dynamic imports, крупном chunk и будущей несовместимости native config loader. Это технический долг, не ошибка выполненной сборки: [vite.config.ts](../../vite.config.ts), [plans.ts](../../src/excel/plans.ts).

## 3. Фактическая архитектура

`Excel → React taskpane / agent loop → локальный Express HTTPS → провайдер модели → SSE → агент → подтверждение → Excel.run → сверка`.

Backend **не выполняет Excel-инструменты**. Выполнение, схемы и подтверждения находятся в панели. `/api/chat` получает от клиента messages/tools, проверяет provider/model и пересылает запрос; учётная запись SaaS не проверяется. В web-режиме панель напрямую обращается к DeepSeek/OpenRouter, минуя локальный сервер: [loop.ts](../../src/agent/loop.ts), [server.ts](../../server/src/server.ts), [webProvider.ts](../../src/taskpane/api/webProvider.ts), [panelMode.ts](../../src/taskpane/panelMode.ts).

### Excel-инструменты

В `TOOL_SPECS` **72 инструмента: 52 mutating, 3 с sideEffect, 17 прочих**. `enable_tools` — дополнительный механизм подключения групп; он не включён в эти 72 записи. Числа получены импортом `TOOL_SPECS` через `node --import tsx`. Допуск зависит от ExcelApi и режима. Источники: [toolSchemas.ts](../../src/excel/toolSchemas.ts), [реестр plans](../../src/excel/plans.ts), [тесты полноты реестра](../../src/excel/toolSchemas.test.ts).

| Группа | Существующие инструменты / проверяемые модули |
| --- | --- |
| Контекст и анализ | `get_active_context`, `list_sheets`, `get_sheet_overview`, `get_range_values`, `search_workbook`, `get_range_details`, `get_conditional_formats`, `audit_workbook`, `profile_range`, `analyze_range`, `recall_snapshot`; [excelTools](../../src/excel/excelTools.ts), [audit](../../src/excel/audit.ts), [dataAnalysis](../../src/excel/dataAnalysis.ts) |
| Запись, очистка, структура | `set_range_values`, `set_ranges_values`, `fill_range`, `change_case`, `trim_text`, `convert_values`, `remove_duplicates`, `copy_sheet`, `rename_sheet`, `delete_sheet`, `insert_columns`, `delete_columns`, `group_rows_columns`, `insert_rows`, `delete_rows`, `sort_range`, `apply_filter`, `append_ranges`, `unpivot_range`; [plans](../../src/excel/plans.ts), [dataCleaning](../../src/excel/dataCleaning.ts) |
| Оформление и таблицы | `format_range`, `create_sheet`, `freeze_panes`, `add_conditional_format`, `create_table`, `set_page_layout`, `convert_table_to_range`, `set_data_validation`, `move_conditional_format`, `apply_color_convention`; [sheetFormatPlans](../../src/excel/sheetFormatPlans.ts), [tablePlans](../../src/excel/tablePlans.ts) |
| Диаграммы, сводные, дашборды | `edit_chart`, `arrange_charts`, `delete_slicer`, `delete_pivot`, `delete_chart`, `filter_pivots`, `update_pivot`, `refresh_pivot`, `add_slicer`, `set_sheet_view`, `create_pivot_table`, `format_chart`, `create_chart`; [pivotPlans](../../src/excel/pivotPlans.ts), [pivotTools](../../src/excel/pivotTools.ts), [dashboardTools](../../src/excel/dashboardTools.ts) |
| Финансы и бухгалтерия | `build_lbo_model`, `build_dcf_model`, `build_three_statement_model`, `add_comparison`, `add_multiples`, `add_share_growth`, `reconcile_ranges`, `build_rsbu_reports`; [dcf](../../src/excel/dcf.ts), [reconcilePlans](../../src/excel/reconcilePlans.ts), [rsbuPlans](../../src/excel/rsbuPlans.ts) |
| Файлы, web, память, копия | `measure_workbook_export`, `create_workbook_backup`, `remember_preference`, `save_scenario`, `get_scenario`, `list_files`, `read_file`, `web_search`, `read_web_page`, `import_file_layout`, `import_file_table`; [fileTools](../../src/excel/fileTools.ts), [webTools](../../src/excel/webTools.ts), [memoryTools](../../src/excel/memoryTools.ts) |

Перед исполнением: allowlist, JSON/Ajv, preflight, привязка к листу, analysis-only guard, подтверждение. Чтение неназванного листа проходит ReadScope; side effects памяти требуют намерения пользователя; web требует включения. После записи различаются `verified`, `applied`, `unknown`, `failed_before_write`; неподтверждённый итог не должен выдаваться за успех. Блокируются известные опасные формулы DDE/WEBSERVICE/CALL и др. Это не универсальный sandbox Excel: [loop](../../src/agent/loop.ts), [readScope](../../src/agent/readScope.ts), [formulaSafety](../../src/excel/formulaSafety.ts), [guarantees tests](../../src/agent/guarantees.test.ts).

### AI Router

В реестре 9 провайдеров: DeepSeek, Kimi, Z.AI, OpenAI, Anthropic, Grok, OpenRouter, собственный Qwen endpoint, локальная Ollama с обнаружением моделей. Поддерживаются Chat Completions, Responses и Anthropic Messages; выбор совместимого маршрута запоминается в памяти. Есть SSE-перевод, prompt caching, отмена upstream при уходе клиента и повторы некоторых сетевых ошибок: [providers](../../server/src/providers.ts), [openaiRoute](../../server/src/openaiRoute.ts), [responsesApi](../../server/src/responsesApi.ts), [anthropicApi](../../server/src/anthropicApi.ts), [promptCache](../../server/src/promptCache.ts), [slowFetch](../../server/src/slowFetch.ts).

Список моделей — факт конфигурации данного SHA, не подтверждение их сегодняшней доступности и цены. Реальные вызовы всех моделей при C0 не выполнялись. Для SaaS нужно сохранить адаптеры, но отделить credentials/entitlements/лимиты от глобальных env и callback lookup. URL Qwen и localhost Ollama нельзя без проверки переносить в публичный gateway: в облаке localhost означает инфраструктуру сервера.

### Авторизация, ключи и лимиты

| Область | Есть сейчас | Пробел для SaaS |
| --- | --- | --- |
| Доступ API | Loopback socket, Host/Origin allowlist, общий локальный panel token; health без token | Нет user identity, tenant context, RBAC, сессий, лицензий |
| Windows BYOK | DPAPI CurrentUser, очередь/атомарная замена файла, наружу hint; env fallback | Один локальный владелец; не tenant key vault |
| Web BYOK | Ключ в localStorage, прямые provider-запросы | Кража при компрометации origin; нет серверного metering/лицензирования этого маршрута |
| Rate limit | 60 API запросов/мин на IP по умолчанию, Map процесса, JSON до 8 MiB | Нет распределённого user/org quota, concurrency limit или денежного резерва |
| Agent budget | 20 ответов, 5 активных минут; 45 чтений, 8 data writes, 60 format, 40 fresh writes, 24 web | Клиентские лимиты обходятся прямым API-запросом; это не финансовый контроль |
| Excel I/O | Обычно до 20 000 ячеек; details до 500; отдельные операции имеют свои пределы | Нужны измерения на больших книгах и реальных host-версиях |
| Usage | Tokens/cache/cost, latency, attempts, 200 последних метрик в RAM | Нет durable ledger, user/org attribution, reservations, сверки потерь usage |

Источники: [localOnly](../../server/src/localOnly.ts), [panelToken](../../server/src/panelToken.ts), [keyStore](../../server/src/keyStore.ts), [DPAPI](../../server/src/dpapi.ts), [webProvider](../../src/taskpane/api/webProvider.ts), [server](../../server/src/server.ts), [agent limits](../../src/agent/loop.ts), [I/O limits](../../src/excel/excelTools.ts), [MetricsStore](../../server/src/usageMetrics.ts), [UI prices](../../src/taskpane/api/prices.ts). UI prices — приблизительная оценка, не бухгалтерский источник.

### Остальные серверные компоненты и данные

| Компонент | Состояние / следствие |
| --- | --- |
| Backup | Локальный диск; 100 MiB/книга, 4 активные загрузки, срез 4 MiB, TTL загрузки 10 мин, retention 7 дней, общий лимит 1 GiB. Это не облачное восстановление: [backupStore](../../server/src/backupStore.ts), [backupRoutes](../../server/src/backupRoutes.ts), [workbookBackup](../../src/excel/workbookBackup.ts) |
| Undo / snapshots | Память панели, undo depth 25; структурные изменения инвалидируют стек; snapshots до 50 / 2 MiB. Перезапуск теряет стек: [undo](../../src/excel/undo.ts), [snapshotStore](../../src/excel/snapshotStore.ts) |
| История | localStorage, 30 дней, до 20 бесед / 2 MiB; хранит сообщения и данные инструментов; не журнал аудита: [conversationStore](../../src/taskpane/conversationStore.ts) |
| Файлы | CSV/TXT/XLSX/DOCX/PDF и 1С; worker с ограничением 60 сек / heap 1024 MiB; до 5 файлов в RAM на 4 часа. Изоляция worker не равна process/container sandbox: [fileStore](../../server/src/files/fileStore.ts), [parse](../../server/src/files/parse.ts) |
| Web | URL/IP/redirect checks, DNS проверяется при соединении, ограничения времени/объёма; поисковые сервисы: [safeFetch](../../server/src/web/safeFetch.ts), [services](../../server/src/web/services.ts) |
| Память / экспорт | Общий локальный JSON, экспорт Markdown в Documents; нет tenancy: [memoryStore](../../server/src/memoryStore.ts), [conversationExport](../../server/src/conversationExport.ts) |
| Эксплуатация | Локальный HTTPS, ротация app.log, health/update, supervisor, release rollback; не SaaS observability и не fleet management: [server](../../server/src/server.ts), [start-server](../../scripts/start-server.ps1), [release](../../scripts/release.ps1) |

## 4. Платформы: что подтверждено

| Платформа | Найденные свидетельства | Чего они не доказывают |
| --- | --- | --- |
| Windows Excel | [README](../../README.md), [STAGE10_PROGRESS](../development/STAGE10_PROGRESS.md) описывают живые прогоны Office 2021 / ExcelApi 1.14; [live-excel-model](../../scripts/live-excel-model.mjs) использует реальную модель; другие live-excel scripts часто подменяют LLM для воспроизводимости | В C0 реальные Excel-прогоны не повторены; отчёты предыдущих дат не сертификат всех функций текущего SHA |
| Windows installer | [windows-bundle.yml](../../.github/workflows/windows-bundle.yml): тесты, bundle smoke, DPAPI, install/uninstall | Наличие workflow не подтверждает последний успешный run; текущие GitHub Actions результаты в C0 не заверены |
| Excel Online | Web build и manifest валидны; общие инструменты и capability gates | [README](../../README.md) прямо говорит, что настоящий Excel Online не проверялся |
| macOS Excel | [web-panel.yml](../../.github/workflows/web-panel.yml) содержит macOS job проверки install/uninstall команды; отчёт о WebKit в [STAGE10_PROGRESS](../development/STAGE10_PROGRESS.md) | На runner нет Excel; WebKit не Excel for Mac. Подтверждённых E2E на настоящем macOS Excel не найдено |

В web-режиме нет локальных backup/memory/file/web-service возможностей. Это существующая продуктовая разница, а не ошибка сборки: [panelMode](../../src/taskpane/panelMode.ts), [needsLocal guard](../../src/agent/loop.ts). Для SaaS их нельзя объявлять доступными до реализации облачных адаптеров и проверки host API.

## 5. Критические операции и тестовое покрытие

Найдено 66 test-файлов в `src`, 24 в `server/src`; npm перечисляет их явно. **770 passed не означает 100% покрытие** или отсутствие data-loss дефектов.

| Инвариант / операция | Доказательство существующих тестов | Недостающая проверка |
| --- | --- | --- |
| Любая запись через план и confirm; analysis-only | [toolSchemas.test](../../src/excel/toolSchemas.test.ts), [loop.test](../../src/agent/loop.test.ts) | Реальный UI и отказ пользователя на всех host |
| Неверные args/опасные формулы | [a1.test](../../src/excel/a1.test.ts), [formulaSafety.test](../../src/excel/formulaSafety.test.ts) | Расширенный hostile workbook corpus, external refs/UDF |
| Partial write, cancel, wrong workbook | [guarantees.test](../../src/agent/guarantees.test.ts), [groupWrite.test](../../src/excel/groupWrite.test.ts) | Crash/reload между write и read-back на настоящем Excel |
| Stale state и newer manual edits | [workbookContext.test](../../src/excel/workbookContext.test.ts), [workbookEvents.test](../../src/excel/workbookEvents.test.ts), [undo.test](../../src/excel/undo.test.ts) | Coauthoring, несколько taskpanes, конфликт нескольких клиентов |
| Строки/листы/таблицы, сортировка | [rowOps.test](../../src/excel/rowOps.test.ts), [sheetOps.test](../../src/excel/sheetOps.test.ts), [tablePlans.test](../../src/excel/tablePlans.test.ts), [sortFilter.test](../../src/excel/sortFilter.test.ts) | Полное восстановление после необратимых операций |
| Сводные/диаграммы | [pivotPlans.test](../../src/excel/pivotPlans.test.ts), [pivotTools.test](../../src/excel/pivotTools.test.ts), [dashboardTools.test](../../src/excel/dashboardTools.test.ts), [chartPlans.test](../../src/excel/chartPlans.test.ts) | Office-specific crash/regression на Windows/Mac/Online |
| Финансовые расчёты | [threeStatement.test](../../src/excel/threeStatement.test.ts), [dcf.test](../../src/excel/dcf.test.ts), [lbo.test](../../src/excel/lbo.test.ts), [reconcileModel.test](../../src/excel/reconcileModel.test.ts), [rsbuModel.test](../../src/excel/rsbuModel.test.ts) | Проверка смысла допущений и итогового текста LLM человеком |
| Область чтения / injection defense | [readScope.test](../../src/agent/readScope.test.ts), [loop.test](../../src/agent/loop.test.ts), [web.test](../../server/src/web/web.test.ts) | Все источники/модели/side effects; prompt injection не доказанно устранён |
| Ключи, API boundary, backup/file safety | [keyStore.test](../../server/src/keyStore.test.ts), [keyRoutes.test](../../server/src/keyRoutes.test.ts), [panelToken.test](../../server/src/panelToken.test.ts), [localOnly.test](../../server/src/localOnly.test.ts), [backupRoutes.test](../../server/src/backupRoutes.test.ts), [parse.test](../../server/src/files/parse.test.ts) | Tenant isolation, SaaS auth, abuse/load, cloud restore |
| Provider wire formats / usage | [protocol.test](../../server/src/protocol.test.ts), [openaiRoute.test](../../server/src/openaiRoute.test.ts), [responsesApi.test](../../server/src/responsesApi.test.ts), [anthropicApi.test](../../server/src/anthropicApi.test.ts), [usageMetrics.test](../../server/src/usageMetrics.test.ts) | Durable metering, lost usage, double charge, spend under concurrency |

## 6. Повторное использование

«Без изменений» ниже означает перенос исходной логики с сохранением контрактов и тестов, **не** отсутствие проверки зависимостей/host compatibility.

| Решение | Компоненты | Условия |
| --- | --- | --- |
| Использовать без переписывания бизнес-логики | A1/formula utilities, pure model builders DCF/LBO/threeStatement/reconcile/RSBU, схемы инструментов, prepare/execute/read-back plans, guarded undo, capability gates | Ссылки на код/тесты в разделах 3–5; сохранить ограничения отмены и независимость от лицензии |
| Использовать без изменений wire-format ядра | protocol, responsesApi translation, openaiRoute parsing, usage normalization | Добавить снаружи tenant-scoped credentials, лимиты и постоянное хранилище; не переносить глобальные mutable lookups как tenant boundary |
| Сохранить для локальной редакции | DPAPI KeyStore, supervisor, installer, local backup | Это однопользовательские компоненты Windows; не использовать как серверную tenancy |
| Использовать с адаптерами | Taskpane/client transport, provider registry, файлы, web services, memory/export, telemetry | Добавить cloud transport, auth context, privacy/retention и resource quotas; сохранить текущий local/BYOK путь |
| Создать заново как отдельные сервисные модули | Users/org membership, entitlements, usage ledger/reservations, hosted key vault, BillingProvider, persistent operation journal | В текущем snapshot отсутствуют; см. [архитектуру](SAAS_ARCHITECTURE.md) |

Проект содержит [Apache-2.0 LICENSE](../../LICENSE) и [NOTICE](../../NOTICE). Они сохраняются; полный license/SBOM review зависимостей и бренда ещё нужен. Оценка компании $1 млн не подтверждается техническим аудитом: коммерческие показатели, права и traction здесь не исследованы.

## 7. Решение C0

Готовая основа — локальный Excel-agent с планами, подтверждениями и multi-provider transport. Коммерческий multi-tenant SaaS отсутствует. Основные блокеры: tenant/auth boundary, долговечный финансовый ledger и лимиты, секреты в hosted/web режиме, восстановление после partial/crash, неподтверждённые Mac/Online. Подробности и очередность: [риски](RISK_REGISTER.md), [архитектура на утверждение](SAAS_ARCHITECTURE.md), [этапы](ROADMAP.md), [приёмка](ACCEPTANCE_CRITERIA.md).

**C1 не начинается до явного утверждения архитектуры и фактической базовой ветки.** Никакой автоматической публикации, merge, тега или выпуска C0 не предполагает.
