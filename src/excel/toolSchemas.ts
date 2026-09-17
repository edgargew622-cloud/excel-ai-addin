import AjvModule, { type ValidateFunction } from "ajv";

/**
 * Единственный источник правды о том, что модели разрешено делать с книгой.
 * Имя здесь = имя функции в excelTools.ts. Всё, чего нет в этом списке,
 * агентский цикл отклоняет, не доходя до Excel.
 */

export type ToolName =
  | "get_active_context"
  | "list_sheets"
  | "get_sheet_overview"
  | "get_range_values"
  | "search_workbook"
  | "get_range_details"
  | "recall_snapshot"
  | "measure_workbook_export"
  | "create_workbook_backup"
  | "set_range_values"
  | "set_ranges_values"
  | "insert_rows"
  | "delete_rows"
  | "sort_range"
  | "apply_filter"
  | "create_pivot_table"
  | "create_chart"
  | "format_range";

export interface ToolSpec {
  name: ToolName;
  /** Операция меняет книгу. Не каждую мутацию безопасно добавлять в custom undo. */
  mutating: boolean;
  /** Необратимо или затирает данные — панель спросит подтверждение. */
  destructive: boolean;
  description: string;
  parameters: Record<string, unknown>;
}

const sheetProp = {
  type: "string",
  description: "Имя листа. Если не указан, используется лист, активный в начале текущей задачи."
};

const addressProp = {
  type: "string",
  description: "A1-адрес без имени листа (B2:D20, H:H, 1:10) или именованный диапазон. Большие чтения всё равно ограничены."
};

export const TOOL_SPECS: ToolSpec[] = [
  {
    name: "get_active_context",
    mutating: false,
    destructive: false,
    description: "Получить идентичность книги в этой сессии, активный лист, активную ячейку, все выделенные области, время чтения и возможности Excel.",
    parameters: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "list_sheets",
    mutating: false,
    destructive: false,
    description: "Перечислить листы книги: ID, имя, порядок, видимость и защиту. Выдача ограничена и сообщает о неполноте.",
    parameters: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "get_sheet_overview",
    mutating: false,
    destructive: false,
    description: "Получить компактный обзор листа: границы данных, таблицы, имена, диаграммы и сводные без чтения всех ячеек.",
    parameters: {
      type: "object",
      properties: { sheet: sheetProp },
      additionalProperties: false
    }
  },
  {
    name: "get_range_values",
    mutating: false,
    destructive: false,
    description:
      "Прочитать значения, формулы и числовые форматы диапазона. Единственный способ узнать содержимое книги. Читай только то, что нужно для задачи.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        properties: {
          type: "array",
          description: "Какие свойства вернуть. По умолчанию values и formulas.",
          items: { type: "string", enum: ["values", "formulas", "text", "valueTypes", "numberFormat"] },
          uniqueItems: true
        }
      },
      required: ["address"],
      additionalProperties: false
    }
  },
  {
    name: "search_workbook",
    mutating: false,
    destructive: false,
    description:
      "Найти значение или текст формулы в используемых областях книги. Возвращает совпадения, реально проверенные области и continuation для следующей порции.",
    parameters: {
      type: "object",
      properties: {
        query: { type: "string", minLength: 1, maxLength: 500, description: "Искомый текст." },
        sheets: {
          type: "array",
          minItems: 1,
          maxItems: 100,
          uniqueItems: true,
          items: { type: "string", minLength: 1 },
          description: "Имена листов. Если не указаны, проверяются все листы."
        },
        searchIn: { type: "string", enum: ["values", "formulas", "both"], description: "Где искать. По умолчанию both." },
        matchCase: { type: "boolean" },
        wholeCell: { type: "boolean" },
        limit: { type: "integer", minimum: 1, maximum: 200, description: "Число совпадений в одной порции. По умолчанию 50." },
        continuation: { type: "string", description: "Непрозрачный курсор из предыдущего ответа с теми же параметрами." }
      },
      required: ["query"],
      additionalProperties: false
    }
  },
  {
    name: "get_range_details",
    mutating: false,
    destructive: false,
    description:
      "Получить оформление, объединения, правила ввода и защиту небольшой области без изменения книги.",
    parameters: {
      type: "object",
      properties: { sheet: sheetProp, address: addressProp },
      required: ["address"],
      additionalProperties: false
    }
  },
  {
    name: "recall_snapshot",
    mutating: false,
    destructive: false,
    description:
      "Вернуть ранее сохранённый снимок чтения как исторические данные и явно сообщить, свежий он, устарел или был вытеснен. " +
      "Признак sheetNameUnverified означает только то, что имя листа в снимке могло устареть: цель определяется по ID листа, " +
      "на достоверность данных это не влияет и перечитывания само по себе не требует.",
    parameters: {
      type: "object",
      properties: {
        snapshotId: { type: "string", minLength: 1, description: "ID снимка из результата инструмента чтения или плана." }
      },
      required: ["snapshotId"],
      additionalProperties: false
    }
  },
  {
    name: "measure_workbook_export",
    mutating: false,
    destructive: false,
    description:
      "Замерить выгрузку всей книги: доступна ли она на этом Excel, каков размер, сколько срезов и сколько занимает по времени. " +
      "Книгу не меняет и никуда её не отправляет: срезы читаются, считаются и отбрасываются. Нужен для подготовки резервных копий.",
    parameters: {
      type: "object",
      properties: {
        sliceSizeBytes: {
          type: "integer",
          minimum: 1024,
          maximum: 4194304,
          description: "Размер среза в байтах. По умолчанию предел Office — 4194304."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "create_workbook_backup",
    mutating: false,
    destructive: false,
    description:
      "Сохранить резервную копию всей книги на этом компьютере через локальный сервер. Книгу не меняет: выгрузка — чтение. " +
      "Копия снимается из открытой книги, а не из файла на диске, поэтому при несохранённых правках они могут различаться. " +
      "Имя и путь задаёт сервер, исходный формат сохраняется. Копия из неполной выгрузки не публикуется.",
    parameters: {
      type: "object",
      properties: {
        sliceSizeBytes: {
          type: "integer",
          minimum: 1024,
          maximum: 4194304,
          description: "Размер среза в байтах. По умолчанию предел Office — 4194304."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "set_range_values",
    mutating: true,
    destructive: true,
    description:
      "Записать значения или формулы в диапазон. Размер массива values должен точно совпадать с размером диапазона: строк × столбцов.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        values: {
          type: "array",
          description:
            "Двумерный массив строк по строкам диапазона. Каждый вложенный массив — одна строка.",
          items: {
            type: "array",
            items: { type: ["string", "number", "boolean", "null"] }
          }
        },
        isFormula: {
          type: "boolean",
          description:
            "true — содержимое values трактуется как формулы (начинаются со знака равенства). По умолчанию false."
        }
      },
      required: ["address", "values"],
      additionalProperties: false
    }
  },
  {
    name: "set_ranges_values",
    mutating: true,
    destructive: true,
    description:
      "Записать значения или формулы в несколько непересекающихся диапазонов одной книги: один план, одно подтверждение, отдельный статус у каждой операции. " +
      "Пересечения диапазонов отклоняются до записи. Операции выполняются по порядку и не являются единой транзакцией: после сбоя оставшиеся не выполняются, " +
      "а уже выполненные не откатываются. Непересечение адресов не доказывает независимость: если одна запись должна опираться на результат другой, " +
      "разделяйте шаги. Для проверки результата записи используйте отдельное чтение после неё, а не эту группу.",
    parameters: {
      type: "object",
      properties: {
        writes: {
          type: "array",
          minItems: 2,
          maxItems: 20,
          description: "Операции записи в порядке исполнения. Для одной операции используйте set_range_values.",
          items: {
            type: "object",
            properties: {
              sheet: sheetProp,
              address: addressProp,
              values: {
                type: "array",
                description: "Двумерный массив строк диапазона, размер обязан совпадать с диапазоном.",
                items: { type: "array", items: { type: ["string", "number", "boolean", "null"] } }
              },
              isFormula: {
                type: "boolean",
                description: "true — содержимое values трактуется как формулы. По умолчанию false."
              }
            },
            required: ["address", "values"],
            additionalProperties: false
          }
        }
      },
      required: ["writes"],
      additionalProperties: false
    }
  },
  {
    name: "insert_rows",
    mutating: true,
    destructive: true,
    description: "Вставить пустые строки, сдвинув существующие вниз.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        startRow: {
          type: "integer",
          description: "Номер строки, перед которой вставлять. Нумерация с 1, как в интерфейсе Excel.",
          minimum: 1
        },
        count: { type: "integer", description: "Сколько строк вставить.", minimum: 1, maximum: 1000 }
      },
      required: ["startRow", "count"],
      additionalProperties: false
    }
  },
  {
    name: "delete_rows",
    mutating: true,
    destructive: true,
    description: "Удалить строки целиком, сдвинув нижние вверх.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        startRow: { type: "integer", description: "Первая удаляемая строка, нумерация с 1.", minimum: 1 },
        count: { type: "integer", description: "Сколько строк удалить.", minimum: 1, maximum: 1000 }
      },
      required: ["startRow", "count"],
      additionalProperties: false
    }
  },
  {
    name: "sort_range",
    mutating: true,
    destructive: true,
    description:
      "Отсортировать область по одному столбцу. Указывай всю сплошную область данных: сортировка части столбцов перемешивает строки, поэтому такая область отклоняется. " +
      "Если в первой строке заголовки, включай hasHeaders, иначе заголовок уедет в середину данных. Перед операцией показывается предпросмотр первых строк сейчас и после, " +
      "после — проверяется, что ни одна строка не потеряна и не перемешана.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        column: {
          type: "integer",
          description: "Индекс столбца внутри диапазона, отсчёт от 0 для первого столбца диапазона.",
          minimum: 0
        },
        ascending: { type: "boolean", description: "По возрастанию. По умолчанию true." },
        hasHeaders: { type: "boolean", description: "Первая строка диапазона — заголовки." },
        allowPartialRows: {
          type: "boolean",
          description: "Разрешить сортировку области уже сплошного блока данных. Ставь только после явного согласия пользователя: соседние столбцы не переедут вместе со строками."
        }
      },
      required: ["address", "column"],
      additionalProperties: false
    }
  },
  {
    name: "apply_filter",
    mutating: true,
    destructive: true,
    description:
      "Применить автофильтр к области по одному столбцу. Скрывает строки, данные не меняет. На листе может быть только один автофильтр: " +
      "фильтр на другую область заменит уже стоящий вместе с его условиями, а на той же области условие добавится к существующим; предпросмотр покажет, что именно произойдёт. Области внутри таблиц Excel отклоняются — у таблицы свой фильтр.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        column: { type: "integer", description: "Индекс столбца внутри диапазона, отсчёт от 0.", minimum: 0 },
        criteria: {
          type: "string",
          description:
            "Условие: точное значение ('Москва') или сравнение ('>100', '<=0', '<>0'). Для нескольких значений перечисли через |, например 'Москва|Ереван'."
        }
      },
      required: ["address", "column", "criteria"],
      additionalProperties: false
    }
  },
  {
    name: "create_pivot_table",
    mutating: true,
    destructive: true,
    description:
      "Создать сводную таблицу. Источник должен включать строку заголовков. Имена в rows и values — это заголовки столбцов источника.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        sourceAddress: { type: "string", description: "Диапазон источника с заголовками, например A1:E200." },
        destAddress: {
          type: "string",
          description: "Левая верхняя ячейка, куда положить сводную, например H1. Место должно быть свободно."
        },
        destSheet: {
          type: "string",
          description: "Лист назначения. Пусто — тот же лист, что и источник."
        },
        rows: {
          type: "array",
          description: "Заголовки столбцов для строк сводной.",
          items: { type: "string" }
        },
        values: {
          type: "array",
          description: "Заголовки столбцов для области значений. Агрегация по умолчанию — сумма.",
          items: { type: "string" }
        }
      },
      required: ["sourceAddress", "destAddress", "rows", "values"],
      additionalProperties: false
    }
  },
  {
    name: "create_chart",
    mutating: true,
    destructive: false,
    description: "Построить диаграмму по диапазону и положить её на лист.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        chartType: {
          type: "string",
          description: "Тип диаграммы.",
          enum: ["ColumnClustered", "Line", "Pie", "BarClustered", "XYScatter", "Area", "Doughnut"]
        },
        title: { type: "string", description: "Заголовок диаграммы." }
      },
      required: ["address", "chartType"],
      additionalProperties: false
    }
  },
  {
    name: "format_range",
    mutating: true,
    destructive: true,
    description:
      "Изменить оформление диапазона: числовой формат, жирность, цвет заливки. Меняются только указанные свойства, остальное оформление не трогается. " +
      "Перед изменением показывается предпросмотр «сейчас → станет», после — результат сверяется обратным чтением. " +
      "Ручная правка оформления между предпросмотром и подтверждением останавливает операцию. Точная отмена возможна на небольших областях.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        numberFormat: {
          type: "string",
          description: "Числовой формат Excel, например '#,##0.00' или '0%'."
        },
        bold: { type: "boolean" },
        fillColor: { type: "string", description: "Цвет заливки в HEX, например #FFF3CD." }
      },
      required: ["address"],
      additionalProperties: false
    }
  }
];

export const TOOL_BY_NAME = new Map<string, ToolSpec>(TOOL_SPECS.map((t) => [t.name, t]));

/** Формат, который ждёт OpenAI-совместимый /chat/completions.
 * Инструменты, не поддерживаемые текущим Excel requirement set, модели не показываем вовсе.
 */
export function toolsForApi(analysisOnly = false) {
  return TOOL_SPECS.filter((spec) =>
    supported(spec) && (!spec.mutating || (!analysisOnly && writableAtCurrentStage(spec)))
  ).map((t) => ({
    type: "function" as const,
    function: {
      name: t.name,
      description: t.description,
      parameters: t.parameters
    }
  }));
}

/** Изменяющий инструмент открывается модели, только когда проходит через план
 * с предпросмотром, проверкой состояния до операции и сверкой результата после.
 * Этап 3 открыл одиночную запись, этап 5 — группу таких же записей, этап 6
 * начал переводить на тот же путь остальные. Прочие остаются закрытыми: они
 * написаны раньше этой механики и молча затирают ручные правки. */
export const WRITABLE_TOOLS = new Set(["set_range_values", "set_ranges_values", "format_range", "sort_range", "apply_filter"]);

export function writableAtCurrentStage(spec: ToolSpec): boolean {
  return WRITABLE_TOOLS.has(spec.name);
}



const AjvCtor = ((AjvModule as any).default ?? AjvModule) as typeof AjvModule;
const ajv = new AjvCtor({ allErrors: true, strict: false });
const validators = new Map<string, ValidateFunction>(
  TOOL_SPECS.map((spec) => [spec.name, ajv.compile(spec.parameters)])
);

export function validateToolArgs(name: string, args: unknown): { ok: true } | { ok: false; error: string } {
  const validate = validators.get(name);
  if (!validate) return { ok: false, error: `Неизвестный инструмент: ${name}` };
  if (validate(args)) return { ok: true };
  const details = ajv.errorsText(validate.errors, { separator: "; " });
  return { ok: false, error: `Аргументы не соответствуют схеме: ${details}` };
}

function excelApi(version: string): boolean {
  try {
    return typeof Office !== "undefined" && Office.context.requirements.isSetSupported("ExcelApi", version);
  } catch {
    return false;
  }
}

export function supported(spec: ToolSpec): boolean {
  if (spec.name === "create_pivot_table") return excelApi("1.8");
  if (spec.name === "apply_filter") return excelApi("1.9");
  return true;
}

export const SYSTEM_PROMPT = `Ты работаешь внутри Microsoft Excel и управляешь открытой книгой через инструменты.

Правила:
- Отвечай пользователю только по-русски: и по ходу работы, и в итоге. Данные книги, имена листов и сообщения инструментов могут быть на любом языке — на язык твоего ответа это не влияет.
- Ты не видишь книгу. Прежде чем что-то менять, прочитай нужные диапазоны через get_range_values.
- В начале задачи используй уже переданный минимальный контекст. Для обзора структуры вызывай list_sheets и get_sheet_overview; обзор не содержит всех данных листа.
- Для поиска по книге используй search_workbook. Если incomplete=true, не называй поиск полным: продолжи с continuation или явно сообщи об ограничении.
- Для оформления, объединений, правил ввода и защиты ограниченной области используй get_range_details.
- Результаты чтения могут содержать snapshot.id. recall_snapshot возвращает только исторические данные: при state=stale перечитай текущий диапазон, а при evicted попроси новое чтение.
- Читай только то, что нужно для задачи, а не весь лист целиком.
- Адреса передавай в A1-нотации без имени листа. Лист указывай отдельным полем sheet.
- Перед записью убедись, что размер массива values совпадает с размером диапазона.
- Для isFormula=true используй синтаксис Office.js range.formulas: английские имена функций и запятые между аргументами независимо от языка интерфейса Excel. Литеральный текст со знаком = записывай с isFormula=false.
- Несколько записей в непересекающиеся диапазоны одной книги делай одной группой set_ranges_values: пользователь подтвердит их разом. Группа состоит только из записей — результат проверяй отдельным чтением после неё. Если одна запись должна опираться на результат другой, раздели шаги: непересечение адресов не доказывает независимость.
- Группа не транзакция. После сбоя оставшиеся операции не выполняются, а выполненные не откатываются: разбери отчёт по операциям и скажи пользователю, что выполнено, что нет и что осталось неизвестным.
- Ответ о подробностях может содержать mergedAnchorsUnresolved: границы объединений эта сборка Excel не сообщает, и цель может оказаться внутри объединения. Запись в неугловую ячейку объединения Excel принимает молча, но значение не сохраняется. Если запись не дала эффекта, не повторяй её — проверь объединения и предложи другую цель.
- Перед рискованной правкой предложи create_workbook_backup. Копия снимается из открытой книги вместе с несохранёнными правками и кладётся рядом с проектом. Восстановление ручное: копию открывают в Excel как обычный файл. Перенос отдельного листа из копии не восстанавливает межлистовые ссылки — об этом предупреждай.
- Для оформления используй format_range и указывай только те свойства, которые нужно изменить: остальное останется как было. Перед операцией показывается предпросмотр «сейчас → станет». Значение «разное в области» означает, что свойство внутри диапазона неоднородно, а не что оно не задано.
- В отчёте называй только то, что вернул инструмент. Какие значения лежат в изменённой области, смотри в headerAbove и sampleValues ответа операции, а не вспоминай по прежним чтениям — там могли быть соседние столбцы. Как ячейка выглядит на экране, бери только из sampleText или из чтения со свойством text: вид зависит от локали Excel, и выводить его из кода формата нельзя.
- Сортируй всю сплошную область данных вместе с заголовками и выставляй hasHeaders, если первая строка — заголовки. Сортировка части столбцов перемешивает строки; allowPartialRows ставь только после явного согласия пользователя. В отчёте о сортировке опирайся на firstRowsAfter и keyHeader из ответа, а orderNote передавай как есть.
- Фильтр на ту же область добавляет условие к уже стоящим, а фильтр на другую область заменяет прежний целиком. Что произошло, бери из filterChange ответа — new, adds, replacesColumn или replacesFilter — и не утверждай, что фильтр заменён, если это не так. В отчёте называй visibleRowsBefore, visibleRowsAfter и hiddenRowsAfter из ответа, а какие условия действуют — бери из conditionsAfter; areaRows — это размер области, а не видимые строки.
- Меняй ровно ту цель, которую назвал пользователь. Если считаешь, что нужна другая — например, всё объединение вместо одной его ячейки, — сначала объясни и спроси, и только после согласия строй операцию. Каким стал формат, бери из поля actual ответа операции, а не из своего запроса.
- Если данные неоднозначны, задай вопрос пользователю вместо того, чтобы угадывать.
- Если инструмент сообщает executionState=applied или unknown, не повторяй запись. Сначала попроси проверить или перечитать текущий диапазон.
- Разрушительные операции пользователь подтверждает вручную. Если подтверждение отклонено, не повторяй ту же операцию — предложи другой вариант.
- Закончив работу, коротко опиши на русском, что именно изменилось и где.`;
