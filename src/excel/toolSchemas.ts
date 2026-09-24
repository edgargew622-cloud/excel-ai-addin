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
  | "profile_range"
  | "get_conditional_formats"
  | "recall_snapshot"
  | "measure_workbook_export"
  | "create_workbook_backup"
  | "set_range_values"
  | "set_ranges_values"
  | "fill_range"
  | "insert_rows"
  | "delete_rows"
  | "sort_range"
  | "apply_filter"
  | "create_pivot_table"
  | "create_chart"
  | "format_range"
  | "freeze_panes"
  | "add_conditional_format"
  | "create_table"
  | "create_sheet"
  | "trim_text"
  | "convert_values"
  | "remove_duplicates"
  | "rename_sheet"
  | "delete_sheet"
  | "insert_columns"
  | "delete_columns"
  | "group_rows_columns"
  | "set_data_validation"
  | "convert_table_to_range"
  | "move_conditional_format"
  | "apply_color_convention";

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
    name: "get_conditional_formats",
    mutating: false,
    destructive: false,
    description:
      "Правила условного форматирования, задевающие область, в порядке приоритета (1 — самый высокий), без изменения книги: " +
      "вид, область действия, условие и оформление. Номер position нужен для move_conditional_format.",
    parameters: {
      type: "object",
      properties: { sheet: sheetProp, address: addressProp },
      required: ["address"],
      additionalProperties: false
    }
  },
  {
    name: "profile_range",
    mutating: false,
    destructive: false,
    description:
      "Профиль таблицы без изменения книги: по каждому столбцу — типы значений, пустоты, лишние пробелы, числа и даты, записанные текстом, " +
      "неоднозначные даты и числа, коды с ведущими нулями; по строкам — пустые строки и дубликаты. Разделители и порядок даты берутся у книги. " +
      "Начинай с него любую очистку данных. Без address берётся вся занятая область листа.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: { type: "string", description: "Область с шапкой, например A1:F500. Пусто — вся занятая область листа." },
        hasHeaders: { type: "boolean", description: "Первая строка — заголовки. По умолчанию true." }
      },
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
    name: "fill_range",
    mutating: true,
    destructive: true,
    description:
      "Заполнить область одной формулой или одним значением. Для формулы записывается первая ячейка, остальное Excel протягивает сам, " +
      "подстраивая относительные ссылки по строкам и столбцам — как при протяжке за угол. Это правильный способ посчитать столбец: " +
      "не перечисляй тысячи значений через set_range_values. Формулу пиши для первой ячейки области, с английскими именами функций. " +
      "Если у столбцов таблицы разные формулы, передай template — первую строку области, по элементу на столбец: " +
      "каждая формула протянется вниз по своему столбцу. Передаётся или value, или template.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        value: {
          type: ["string", "number", "boolean"],
          description: "Формула для первой ячейки области при isFormula=true, иначе одно значение на все ячейки."
        },
        isFormula: {
          type: "boolean",
          description: "true — value трактуется как формула и протягивается по области. По умолчанию false."
        },
        template: {
          type: "array",
          minItems: 1,
          items: { type: ["string", "number", "boolean"] },
          description: "Первая строка области, по элементу на каждый столбец: формулы начинаются с =, остальное — значения. Заполняется вниз."
        }
      },
      required: ["address"],
      additionalProperties: false
    }
  },
  {
    name: "trim_text",
    mutating: true,
    destructive: true,
    description:
      "Убрать лишние пробелы в тексте области: по краям, неразрывные и (по умолчанию) повторы внутри — как функция TRIM. " +
      "Меняются только текстовые ячейки; формулы и числа не трогаются, текст остаётся текстом. Отменяется кнопкой «Отменить».",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        collapseInner: { type: "boolean", description: "Сжимать повторы пробелов внутри текста. По умолчанию true." }
      },
      required: ["address"],
      additionalProperties: false
    }
  },
  {
    name: "convert_values",
    mutating: true,
    destructive: true,
    description:
      "Превратить числа или даты, записанные текстом, в настоящие числа или даты. Без decimalSeparator и dateOrder меняются только " +
      "однозначные значения по разделителям книги; неоднозначные (1,500; 01.02.2026) и коды с ведущими нулями пропускаются с причиной. " +
      "Даты получают формат даты книги. Отменяется кнопкой «Отменить».",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        to: { type: "string", enum: ["number", "date"] },
        decimalSeparator: { type: "string", enum: [",", "."], description: "Десятичный разделитель — только если его назвал пользователь." },
        dateOrder: { type: "string", enum: ["DMY", "MDY", "YMD"], description: "Порядок даты — только если его назвал пользователь." }
      },
      required: ["address", "to"],
      additionalProperties: false
    }
  },
  {
    name: "remove_duplicates",
    mutating: true,
    destructive: true,
    description:
      "Удалить повторяющиеся строки таблицы: остаётся первое вхождение, строки ниже поднимаются внутри области. Отмены нет — " +
      "перед этим предложи create_workbook_backup. Область — вся таблица с шапкой: если рядом есть данные, в формулах области или она " +
      "задевает таблицу Excel, операция откажет. Предпросмотр называет удаляемые строки и формулы книги, которые станут смотреть на другие данные.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        columns: { type: "array", items: { type: "string" }, description: "Ключ: заголовки или буквы столбцов, по которым строки одинаковы. Пусто — все столбцы." },
        destSheet: { type: "string", description: "Пустой лист для результата: источник не меняется, уникальные строки копируются туда значениями, и есть отмена. Лист создай через create_sheet." },
        hasHeaders: { type: "boolean", description: "Первая строка — заголовки. По умолчанию true." }
      },
      required: ["address"],
      additionalProperties: false
    }
  },
  {
    name: "rename_sheet",
    mutating: true,
    destructive: true,
    description:
      "Переименовать лист. Ссылки на него Excel перепишет сам; предпросмотр назовёт формулы, где имя стоит внутри текста (INDIRECT, HYPERLINK), — " +
      "их Excel не перепишет. Отменяется кнопкой «Отменить».",
    parameters: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "Текущее имя листа." },
        newName: { type: "string", description: "Новое имя: до 31 знака, без : \\ / ? * [ ]." }
      },
      required: ["sheet", "newName"],
      additionalProperties: false
    }
  },
  {
    name: "delete_sheet",
    mutating: true,
    destructive: true,
    description:
      "Удалить лист со всеми данными, диаграммами, сводными и таблицами. Отмены нет — перед удалением листа с данными предложи create_workbook_backup. " +
      "Ссылки других листов на удалённый станут #ССЫЛКА! — предпросмотр их перечислит. Последний видимый лист удалить нельзя.",
    parameters: {
      type: "object",
      properties: { sheet: { type: "string", description: "Имя удаляемого листа." } },
      required: ["sheet"],
      additionalProperties: false
    }
  },
  {
    name: "insert_columns",
    mutating: true,
    destructive: true,
    description:
      "Вставить пустые столбцы, сдвинув существующие вправо. Отката нет — история отмены очищается. Предпросмотр называет формулы, " +
      "которые не охватят новые столбцы (итог по диапазону, вплотную к которому идёт вставка); после операции сверяются ошибки ссылок в книге.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        startColumn: { type: "string", description: "Буквы столбца, например C: вставка — перед ним, удаление — начиная с него." },
        count: { type: "integer", description: "Сколько столбцов.", minimum: 1, maximum: 100 }
      },
      required: ["startColumn", "count"],
      additionalProperties: false
    }
  },
  {
    name: "delete_columns",
    mutating: true,
    destructive: true,
    description:
      "Удалить столбцы, сдвинув остальные влево. Отката нет — перед удалением столбцов с данными предложи create_workbook_backup. " +
      "Предпросмотр называет, какие формулы книги станут #ССЫЛКА! и какие диапазоны молча сузятся; после операции сверяются ошибки ссылок.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        startColumn: { type: "string", description: "Буквы столбца, например C: вставка — перед ним, удаление — начиная с него." },
        count: { type: "integer", description: "Сколько столбцов.", minimum: 1, maximum: 100 }
      },
      required: ["startColumn", "count"],
      additionalProperties: false
    }
  },
  {
    name: "group_rows_columns",
    mutating: true,
    destructive: true,
    description:
      "Сгруппировать целые строки (address «3:10») или столбцы («C:F») — структура Excel с кнопками «+» и «−», а не скрытие. " +
      "collapse=true оставит группу свёрнутой. Прежние уровни группировки Excel через API не сообщает. Отменяется кнопкой «Отменить».",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: { type: "string", description: "Целые строки «3:10» или столбцы «C:F»." },
        collapse: { type: "boolean", description: "Свернуть группу после создания. По умолчанию false." }
      },
      required: ["address"],
      additionalProperties: false
    }
  },
  {
    name: "apply_color_convention",
    mutating: true,
    destructive: true,
    description:
      "Раскрасить текст модели по роли ячеек: введённые числа (input), формулы на этом листе (formula), формулы со ссылкой на другой лист или книгу (link), " +
      "контрольные ячейки из checks (check). Роль панель определяет сама по содержимому; подписи и пустые ячейки не трогаются. " +
      "Палитра — выбор пользователя; по умолчанию синий, чёрный, зелёный, тёмно-красный. Отменяется кнопкой «Отменить».",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        checks: { type: "string", description: "Контрольные ячейки внутри address, например B20:F20." },
        palette: {
          type: "object",
          properties: {
            input: { type: "string" },
            formula: { type: "string" },
            link: { type: "string" },
            check: { type: "string" }
          },
          additionalProperties: false,
          description: "Цвета текста в HEX; незаданные — по умолчанию."
        }
      },
      required: ["address"],
      additionalProperties: false
    }
  },
  {
    name: "move_conditional_format",
    mutating: true,
    destructive: true,
    description:
      "Поднять правило условного форматирования выше всех правил области (to: first) или опустить ниже всех (to: last). " +
      "position — номер правила из get_conditional_formats для той же области. Отменяется кнопкой «Отменить».",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        position: { type: "integer", minimum: 1 },
        to: { type: "string", enum: ["first", "last"] }
      },
      required: ["address", "position", "to"],
      additionalProperties: false
    }
  },
  {
    name: "convert_table_to_range",
    mutating: true,
    destructive: true,
    description:
      "Превратить таблицу Excel в обычный диапазон. Данные остаются, оформление стиля остаётся на ячейках, фильтр таблицы снимается, " +
      "ссылки на таблицу Excel переписывает в обычные адреса. Укажите table (имя) или address ячейки внутри таблицы. Отмены нет.",
    parameters: {
      type: "object",
      properties: {
        table: { type: "string", description: "Имя таблицы Excel." },
        sheet: sheetProp,
        address: { type: "string", description: "Ячейка внутри таблицы, если имя неизвестно." }
      },
      additionalProperties: false
    }
  },
  {
    name: "set_data_validation",
    mutating: true,
    destructive: true,
    description:
      "Поставить на область правило проверки ввода: список допустимых значений, целое число, число или дату в пределах. Заменяет прежнее правило " +
      "области. Уже введённое не меняется: ответ назовёт ячейки, которые правилу не соответствуют. Отменяется кнопкой «Отменить».",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        rule: { type: "string", enum: ["list", "wholeNumber", "decimal", "date"] },
        items: { type: "array", items: { type: "string" }, description: "Для list — допустимые значения, без запятых внутри." },
        operator: { type: "string", enum: ["between", "notBetween", "equalTo", "notEqualTo", "greaterThan", "lessThan", "greaterOrEqual", "lessOrEqual"], description: "Для чисел и дат. По умолчанию between." },
        value: { type: ["number", "string"], description: "Число или дата ГГГГ-ММ-ДД." },
        value2: { type: ["number", "string"], description: "Верхняя граница для between и notBetween." }
      },
      required: ["address", "rule"],
      additionalProperties: false
    }
  },
  {
    name: "insert_rows",
    mutating: true,
    destructive: true,
    description:
      "Вставить пустые строки, сдвинув существующие вниз. Отката нет ни у панели, ни у истории отмены — она очищается целиком. " +
      "Перед операцией показывается предпросмотр: сколько строк сдвинется и какие формулы книги не охватят новые строки " +
      "(например, СУММ по диапазону, вплотную к которому идёт вставка). После операции сверяется число ошибок ссылок в книге. " +
      "Перед вставкой в книгу с данными предлагай create_workbook_backup.",
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
    description:
      "Удалить строки целиком, сдвинув нижние вверх. Это необратимо: ни панель, ни повтор операции строки не вернут. " +
      "Предпросмотр показывает содержимое удаляемых строк, число непустых ячеек и формулы книги, которые станут #ССЫЛКА! " +
      "или молча пересчитаются по укороченному диапазону. После операции сверяется число ошибок ссылок. " +
      "Обязательно предлагай create_workbook_backup перед удалением и не выполняй его, пока пользователь не ответил.",
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
      "Создать сводную таблицу. Источник — вся область с шапкой; имена в rows и values — заголовки её столбцов. " +
      "Панель заранее сама считает сводную: группы, итоги и размер. Если место под сводной занято, операция отклоняется, а не затирает данные. " +
      "После построения итоги сверяются с расчётом панели. По умолчанию сводная встаёт правее данных; для отдельного листа укажи destSheet.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        sourceAddress: { type: "string", description: "Область источника вместе с шапкой, например A1:E200." },
        destAddress: { type: "string", description: "Левая верхняя ячейка сводной, например H1. По умолчанию — через столбец правее данных листа назначения." },
        destSheet: { type: "string", description: "Существующий лист назначения — только если пользователь прямо попросил отдельный лист. Пусто — тот же лист, что и источник." },
        newSheet: { type: "string", description: "Имя нового листа под сводную — если пользователь просит сводную на новом листе. Лист создастся этой же операцией, сводная встанет в A1; отмена уберёт и сводную, и пустой лист." },
        rows: {
          type: "array",
          minItems: 1,
          description: "Заголовки столбцов для строк сводной, от внешнего к внутреннему.",
          items: { type: "string" }
        },
        values: {
          type: "array",
          minItems: 1,
          description: "Поля значений: заголовок столбца и агрегация. Сумма и среднее имеют смысл только для чисел; для текста — count.",
          items: {
            type: "object",
            properties: {
              field: { type: "string" },
              aggregation: { type: "string", enum: ["sum", "count", "average", "max", "min"] }
            },
            required: ["field"],
            additionalProperties: false
          }
        }
      },
      required: ["sourceAddress", "rows", "values"],
      additionalProperties: false
    }
  },
  {
    name: "create_chart",
    mutating: true,
    destructive: true,
    description:
      "Построить диаграмму по области и положить её на лист правее данных, чтобы не закрыть их. " +
      "Предпросмотр называет ряды, число точек и подписи, которые должны получиться; после построения ряды сверяются с тем, что сообщил Excel. " +
      "Если Excel понял область иначе (например, шапку как ряд), это названо, а диаграмму можно убрать отменой.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: { ...addressProp, description: "Область данных вместе с шапкой и столбцом подписей, например A1:C7." },
        chartType: {
          type: "string",
          description: "Тип диаграммы: ColumnClustered — столбцы, BarClustered — полосы, Line — линии, Area — области, Pie — круговая (один ряд), Doughnut — кольцевая, XYScatter — точечная.",
          enum: ["ColumnClustered", "Line", "Pie", "BarClustered", "XYScatter", "Area", "Doughnut"]
        },
        title: { type: "string", description: "Заголовок диаграммы." },
        seriesBy: {
          type: "string",
          enum: ["columns", "rows"],
          description: "columns (по умолчанию) — каждый столбец ряд, строки — точки; rows — наоборот."
        },
        anchorCell: {
          type: "string",
          description: "Ячейка левого верхнего угла диаграммы, например H2. По умолчанию — через столбец правее занятой области листа."
        }
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
      "Изменить оформление диапазона: числовой формат, шрифт, заливку, выравнивание, перенос, границы, ширину столбцов и высоту строк. " +
      "Меняются только указанные свойства, остальное оформление не трогается. " +
      "Перед изменением показывается предпросмотр «сейчас → станет», после — результат сверяется обратным чтением. " +
      "Ручная правка оформления между предпросмотром и подтверждением останавливает операцию. Точная отмена возможна на областях до 500 ячеек. " +
      "Чтобы оформить таблицу целиком, обычно нужны два вызова: шапка (жирный, заливка, выравнивание) и вся область (границы, автоподбор ширины).",
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
        italic: { type: "boolean" },
        underline: { type: "boolean", description: "true — одинарное подчёркивание, false — без подчёркивания." },
        fontColor: { type: "string", description: "Цвет текста в HEX, например #1F4E79." },
        fontSize: { type: "number", minimum: 1, maximum: 409, description: "Размер шрифта в пунктах." },
        fontName: { type: "string", description: "Имя шрифта, например Calibri или Arial." },
        fillColor: { type: "string", description: "Цвет заливки в HEX, например #FFF3CD." },
        horizontalAlignment: { type: "string", enum: ["General", "Left", "Center", "Right", "Justify"] },
        verticalAlignment: { type: "string", enum: ["Top", "Center", "Bottom"] },
        wrapText: { type: "boolean", description: "Перенос текста по словам внутри ячейки." },
        borders: {
          type: "string",
          enum: ["all", "outline", "inside", "none"],
          description: "all — сетка по всей области, outline — только внешняя рамка, inside — только внутренние линии, none — убрать все границы области."
        },
        borderColor: { type: "string", description: "Цвет линий в HEX, по умолчанию #000000. Только вместе с borders." },
        borderWeight: { type: "string", enum: ["Thin", "Medium", "Thick"], description: "Толщина линий, по умолчанию Thin. Только вместе с borders." },
        columnWidth: {
          type: "number",
          minimum: 0,
          maximum: 1600,
          description: "Ширина каждого столбца области в пунктах. Обычно удобнее columnWidthChars в знаках; пункты — только если пользователь сам назвал их."
        },
        columnWidthChars: {
          type: "number",
          minimum: 0,
          maximum: 255,
          description: "Ширина каждого столбца области в знаках — как её показывает Excel в интерфейсе. Предпочтительнее columnWidth: панель переведёт знаки в пункты по шрифту самой книги. Не вместе с columnWidth."
        },
        rowHeight: { type: "number", minimum: 0, maximum: 409, description: "Высота каждой строки области в пунктах." },
        autofit: {
          type: "string",
          enum: ["columns", "rows", "both"],
          description: "Автоподбор ширины столбцов и/или высоты строк по содержимому. Нельзя сочетать с явным columnWidth или rowHeight для того же измерения."
        }
      },
      required: ["address"],
      additionalProperties: false
    }
  },
  {
    name: "create_sheet",
    mutating: true,
    destructive: true,
    description:
      "Создать новый пустой лист в этой книге. Имя проверяется заранее по правилам Excel: не длиннее 31 знака, без : \\ / ? * [ ], не повторяет имя существующего листа. " +
      "Данные и другие листы не меняются. Отмена удаляет созданный лист, но только пока он пуст.",
    parameters: {
      type: "object",
      properties: {
        name: { type: "string", description: "Имя нового листа, например «Сводка»." },
        after: { type: "string", description: "Имя листа, сразу после которого поставить новый. Пусто — новый лист встанет последним." }
      },
      required: ["name"],
      additionalProperties: false
    }
  },
  {
    name: "freeze_panes",
    mutating: true,
    destructive: true,
    description:
      "Закрепить верхние строки и/или левые столбцы листа, чтобы они не уходили при прокрутке, или снять закрепление (rows=0 и columns=0). " +
      "Это настройка вида листа: данные и оформление ячеек не меняются. Результат сверяется по месту закрепления, отмена возможна.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        rows: { type: "integer", minimum: 0, maximum: 100, description: "Сколько верхних строк закрепить. Для шапки таблицы в первой строке — 1." },
        columns: { type: "integer", minimum: 0, maximum: 50, description: "Сколько левых столбцов закрепить." }
      },
      additionalProperties: false
    }
  },
  {
    name: "add_conditional_format",
    mutating: true,
    destructive: true,
    description:
      "Добавить правило условного форматирования: подсветка по значению (greaterThan, lessThan, greaterOrEqual, lessOrEqual, equalTo, notEqualTo, between), " +
      "по тексту (textContains), по своей формуле (formula), цветовая шкала (colorScale) или гистограмма в ячейках (dataBar). " +
      "Правило добавляется к уже существующим и не заменяет их; по умолчанию встаёт выше них (order: first), как в Excel. " +
      "Предпросмотр показывает оценку, сколько ячеек подсветится; после операции правило сверяется обратным чтением. Отмена удаляет добавленное правило.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: addressProp,
        rule: {
          type: "string",
          enum: ["greaterThan", "lessThan", "greaterOrEqual", "lessOrEqual", "equalTo", "notEqualTo", "between", "textContains", "formula", "colorScale", "dataBar"]
        },
        formula: {
          type: "string",
          description: "formula: условие для левой верхней ячейки области, английские имена функций. Столбец закрепляйте $: =$C2>1000 подсветит всю строку."
        },
        order: { type: "string", enum: ["first", "last"], description: "first (по умолчанию) — выше прежних правил области, last — ниже них." },
        value: {
          type: ["number", "string"],
          description: "С чем сравнивать. Для between — нижняя граница. Текст допустим только для equalTo и notEqualTo."
        },
        value2: { type: "number", description: "Верхняя граница для between." },
        text: { type: "string", description: "Искомый текст для textContains; регистр не различается." },
        fillColor: { type: "string", description: "Заливка подсвеченных ячеек в HEX." },
        fontColor: { type: "string", description: "Цвет текста подсвеченных ячеек в HEX." },
        bold: { type: "boolean", description: "Жирный текст подсвеченных ячеек." },
        minColor: { type: "string", description: "colorScale: цвет наименьшего значения в HEX." },
        midColor: { type: "string", description: "colorScale: цвет середины (50-й процентиль) в HEX, по желанию." },
        maxColor: { type: "string", description: "colorScale: цвет наибольшего значения в HEX." },
        barColor: { type: "string", description: "dataBar: цвет полосы в HEX, по умолчанию #638EC6." }
      },
      required: ["address", "rule"],
      additionalProperties: false
    }
  },
  {
    name: "create_table",
    mutating: true,
    destructive: true,
    description:
      "Превратить область в таблицу Excel со стилем (полосы строк, шапка с фильтром). Первая строка области обязана быть заголовками. " +
      "Это меняет поведение области: запись вплотную расширяет таблицу, формула в столбце протягивается на весь столбец. " +
      "Отмены в панели нет. Предпросмотр предупреждает, какие заголовки Excel переименует, и отказывает при пересечении с другой таблицей или объединениях.",
    parameters: {
      type: "object",
      properties: {
        sheet: sheetProp,
        address: { ...addressProp, description: "Вся область вместе со строкой заголовков, например A1:E7." },
        style: {
          type: "string",
          description: "Встроенный стиль: TableStyleLight1–21, TableStyleMedium1–28, TableStyleDark1–11. По умолчанию TableStyleMedium2."
        },
        name: { type: "string", description: "Имя таблицы латиницей или кириллицей без пробелов, например Сотрудники. По желанию." }
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
export const WRITABLE_TOOLS = new Set([
  "set_range_values",
  "set_ranges_values",
  "fill_range",
  "format_range",
  "sort_range",
  "apply_filter",
  "insert_rows",
  "delete_rows",
  "freeze_panes",
  "add_conditional_format",
  "create_table",
  "create_chart",
  "create_pivot_table",
  "create_sheet",
  "trim_text",
  "convert_values",
  "remove_duplicates",
  "rename_sheet",
  "delete_sheet",
  "insert_columns",
  "delete_columns",
  "group_rows_columns",
  "set_data_validation",
  "convert_table_to_range",
  "move_conditional_format",
  "apply_color_convention"
]);

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

/**
 * Минимальный набор ExcelApi каждого инструмента (этап 7, 7.1.6).
 *
 * Прежде проверялись четыре инструмента, перечисленные по именам, а прочие
 * выдавались модели без проверки. Теперь набор объявлен у всех, и тест не
 * даёт добавить инструмент без него. Указан набор того, что используется
 * без запасного пути: у объединений (1.13), правил ввода (1.8), выделения
 * (1.9), активной диаграммы есть обход, и они сюда не входят. Защита листа
 * и ячеек — 1.2. Проверочная машина держит 1.14: выше ставить нельзя.
 */
export const MIN_EXCEL_API: Record<ToolName, string> = {
  get_active_context: "1.1",
  list_sheets: "1.2",
  get_sheet_overview: "1.2",
  get_range_values: "1.1",
  search_workbook: "1.1",
  get_range_details: "1.2",
  // Разделители и порядок даты книги — cultureInfo, ExcelApi 1.12.
  profile_range: "1.12",
  recall_snapshot: "1.1",
  measure_workbook_export: "1.1",
  create_workbook_backup: "1.1",
  set_range_values: "1.2",
  set_ranges_values: "1.2",
  fill_range: "1.2",
  insert_rows: "1.2",
  delete_rows: "1.2",
  sort_range: "1.2",
  apply_filter: "1.9",
  create_pivot_table: "1.8",
  create_chart: "1.2",
  format_range: "1.2",
  create_sheet: "1.1",
  freeze_panes: "1.7",
  add_conditional_format: "1.6",
  create_table: "1.2",
  trim_text: "1.2",
  // Разделители и шаблон даты книги — cultureInfo, ExcelApi 1.12.
  convert_values: "1.12",
  remove_duplicates: "1.9",
  // Защита структуры книги и формулы именованных диапазонов — ExcelApi 1.7.
  rename_sheet: "1.7",
  delete_sheet: "1.7",
  insert_columns: "1.2",
  delete_columns: "1.2",
  // Range.group, hideGroupDetails — ExcelApi 1.10.
  group_rows_columns: "1.10",
  // DataValidation — 1.8, getInvalidCellsOrNullObject — 1.9.
  set_data_validation: "1.9",
  convert_table_to_range: "1.2",
  get_conditional_formats: "1.6",
  move_conditional_format: "1.6",
  apply_color_convention: "1.2"
};

export function supported(spec: ToolSpec): boolean {
  return excelApi(MIN_EXCEL_API[spec.name]);
}

export const SYSTEM_PROMPT = `Ты работаешь внутри Microsoft Excel и управляешь открытой книгой через инструменты.

Правила:
- Отвечай пользователю только по-русски: и по ходу работы, и в итоге. Данные книги, имена листов и сообщения инструментов могут быть на любом языке — на язык твоего ответа это не влияет.
- Ты не видишь книгу. Прежде чем что-то менять, прочитай нужные диапазоны через get_range_values.
- В начале задачи используй уже переданный минимальный контекст. Для обзора структуры вызывай list_sheets и get_sheet_overview; обзор не содержит всех данных листа.
- Для поиска по книге используй search_workbook. Если incomplete=true, не называй поиск полным: продолжи с continuation или явно сообщи об ограничении.
- Для оформления, объединений, правил ввода и защиты ограниченной области используй get_range_details.
- Лишние пробелы убирай через trim_text: текст остаётся текстом, коды с нулями не страдают. Числа и даты, записанные текстом, превращай в настоящие через convert_values: по умолчанию меняются только однозначные значения по разделителям книги; decimalSeparator и dateOrder передавай, только когда пользователь их назвал. Предпросмотр показывает пары «было → станет» и пропуски с причинами — перескажи пропуски пользователю.
- Дубликаты строк удаляй через remove_duplicates по всей таблице с шапкой. На месте отмены нет: сначала предложи пользователю выбор — резервную копию книги (create_workbook_backup) или результат на отдельном пустом листе (destSheet, лист — через create_sheet), где источник не меняется и отмена есть, — и дождись ответа. Ключ — столбцы, по которым строки считаются одинаковыми; без него — все столбцы. Excel не различает регистр, но различает пробел в конце: если дубликаты не нашлись из-за пробелов, отказ это скажет — предложи сначала trim_text. Если в ответе есть affectedFormulas, перечисли их: эти формулы теперь смотрят на другие строки.
- Очистку данных начинай с profile_range: он показывает, что мешает считать — числа и даты текстом, лишние пробелы, дубликаты. Неоднозначные даты (01.02.2026) и числа (1,500) не преобразуй без ответа пользователя: спроси, какой порядок или разделитель имелся в виду. Коды с ведущими нулями — не числа. Если incomplete=true, говори только о проверенной области.
- Результаты чтения могут содержать snapshot.id. recall_snapshot возвращает только исторические данные: при state=stale перечитай текущий диапазон, а при evicted попроси новое чтение.
- Читай только то, что нужно для задачи, а не весь лист целиком.
- Адреса передавай в A1-нотации без имени листа. Лист указывай отдельным полем sheet.
- Перед записью убедись, что размер массива values совпадает с размером диапазона.
- Для isFormula=true используй синтаксис Office.js range.formulas: английские имена функций и запятые между аргументами независимо от языка интерфейса Excel. Литеральный текст со знаком = записывай с isFormula=false.
- Запись вплотную к таблице Excel расширяет её: Excel добавляет столбец или строку с автоматическим заголовком. Предупреждение об этом есть в предпросмотре, а по факту изменение приходит в tableChanges ответа — если оно там есть, обязательно скажи пользователю, что границы таблицы изменились.
- Если пользователь просит посчитать и записать результат в книгу, записывай формулы (fill_range, set_range_values с isFormula=true), а не числа, посчитанные тобой: результат должен считать Excel и пересчитывать при изменении данных. Значения вместо формул — только если пользователь прямо просит «значениями», «без формул» или «зафиксировать». Число, которое ты посчитал сам, называй своим расчётом; результатом Excel называй только то, что вернул инструмент после записи.
- Имена функций пиши только по-английски. Панель до записи проверяет, есть ли функции формулы в этом Excel: набор зависит от версии. Если ответ говорит, что функции нет, не подменяй её молча — назови ограничение и предложи пользователю замену явно. Если в ответе есть errorNote или newErrors, обязательно назови ячейки с ошибками и причину: запись выполнена, но результат — ошибка.
- Одну и ту же формулу или одно значение на всю область записывай через fill_range: формулу для первой ячейки Excel протянет сам, подстроив ссылки. Перечислять тысячи значений в set_range_values нельзя — ответ модели упрётся в предел длины, и запись не состоится.
- Несколько записей в непересекающиеся диапазоны одной книги делай одной группой set_ranges_values: пользователь подтвердит их разом. Группа состоит только из записей — результат проверяй отдельным чтением после неё. Если одна запись должна опираться на результат другой, раздели шаги: непересечение адресов не доказывает независимость.
- Группа не транзакция. После сбоя оставшиеся операции не выполняются, а выполненные не откатываются: разбери отчёт по операциям и скажи пользователю, что выполнено, что нет и что осталось неизвестным.
- Ответ о подробностях может содержать mergedAnchorsUnresolved: границы объединений эта сборка Excel не сообщает, и цель может оказаться внутри объединения. Запись в неугловую ячейку объединения Excel принимает молча, но значение не сохраняется. Если запись не дала эффекта, не повторяй её — проверь объединения и предложи другую цель.
- Перед рискованной правкой предложи create_workbook_backup. Копия снимается из открытой книги вместе с несохранёнными правками и кладётся рядом с проектом. Восстановление ручное: копию открывают в Excel как обычный файл. Перенос отдельного листа из копии не восстанавливает межлистовые ссылки — об этом предупреждай.
- Для оформления используй format_range и указывай только те свойства, которые нужно изменить: остальное останется как было. Цвета передавай в HEX (#RRGGBB). Ширину столбцов задавай в знаках через columnWidthChars — так её понимает человек и показывает Excel; соотношение знаков и пунктов не вычисляй сам, его меряет панель по книге. В отчёте называй ширину в знаках из columnWidthChars и widthChars ответа. Высоту строк задавай в пунктах. Ширину и высоту можно менять и для целых столбцов и строк (A:E, 1:10), остальное оформление — только для области данных. Итог автоподбора заранее неизвестен: фактические размеры бери из sizesAfter ответа. Границы и размеры Excel может слегка подогнать к сетке экрана — каким стало свойство, смотри в actual. Перед операцией показывается предпросмотр «сейчас → станет». Значение «разное в области» означает, что свойство внутри диапазона неоднородно, а не что оно не задано.
- В отчёте называй только то, что вернул инструмент. Какие значения лежат в изменённой области, смотри в headerAbove и sampleValues ответа операции, а не вспоминай по прежним чтениям — там могли быть соседние столбцы. Как ячейка выглядит на экране, бери только из sampleText или из чтения со свойством text: вид зависит от локали Excel, и выводить его из кода формата нельзя.
- Если правило условного форматирования не видно, потому что его перекрывает другое, прочитай порядок через get_conditional_formats и перенеси нужное через move_conditional_format; номер position бери из того же списка той же области.
- Цвета финансовой модели ставь через apply_color_convention: роли ячеек определяет панель по содержимому, не перечисляй ячейки сам. Если пользователь назвал свои цвета — передай palette; иначе скажи, что взята палитра по умолчанию. Контрольные строки передавай в checks. Если в ответе есть conditionalNote или overwritten — перескажи их.
- Таблицу Excel в обычный диапазон переводи через convert_table_to_range — только по прямой просьбе. Оформление стиля останется на ячейках: скажи об этом. Ссылки на таблицу Excel перепишет в обычные адреса; если в ответе есть brokenFormulas или filterNote — перескажи их.
- Правила проверки ввода ставь через set_data_validation: список, целое число, число, дата. Правило не исправляет уже введённое — если в ответе есть invalidExistingCells, назови эти ячейки пользователю и не обещай их автоматической очистки. Даты в правиле — ГГГГ-ММ-ДД. В списке регистр важен: «в работе» не пройдёт при «В работе».
- Группируй строки и столбцы через group_rows_columns — это настоящая структура Excel с кнопками «+» и «−»; не подменяй её скрытием строк. Снимать чужие группы агент пока не умеет; об этом говори прямо.
- Столбцы вставляй и удаляй через insert_columns и delete_columns: правила те же, что у строк, — отмены нет, перед удалением столбцов с данными сначала предложи create_workbook_backup и дождись ответа, а affectedFormulas перечисли целиком. Столбцы через таблицу Excel не поддержаны — отказ это скажет.
- Вставка и удаление строк необратимы: отмены в панели нет, история отмены очищается, повтор операции ничего не вернёт. Перед удалением строк с данными сначала предложи create_workbook_backup и дождись ответа пользователя — не выполняй удаление в том же шаге.
- Предпросмотр этих операций перечисляет формулы книги, которые пострадают: broken — станет #ССЫЛКА!, shrunk — диапазон молча уменьшится, missed — существующая формула не охватит вставленные строки. Прежде чем звать инструмент, назови пользователю эти последствия своими словами.
- Разбор формул неполон: структурированные ссылки таблиц (Таблица[Столбец]) не разбираются, а слишком большие листы не обходятся вовсе. Если ответ содержит unscannedSheets или предпросмотр назвал листы с табличными ссылками, скажи прямо, что про них ничего не проверено, и не выдавай проверку за полную.
- После операции сверь refErrorsBefore и refErrorsAfter. Если пришло newRefErrors, обязательно перечисли brokenCells и скажи, что в книге появились сломанные ссылки — сами по себе они не исчезнут.
- Предпросмотр этих операций показывают пользователю, а не тебе: до вызова инструмента ты не знаешь, какие формулы пострадают, и не выдавай догадку за разбор. Разбор приходит в ответе полем affectedFormulas — перечисли его целиком, с листами и адресами. Молчать о нём нельзя даже тогда, когда целевые ячейки выглядят правильно: у вставки такие формулы не дают ни ошибки, ни внешнего признака, и кроме тебя о них никто не скажет.
- Новый лист создаётся через create_sheet: он пуст, данные и другие листы не трогаются. Создавай его, когда пользователь просит отдельный лист или когда результат некуда положить, — но не по своей инициативе ради «чистого места». Переименовывай лист через rename_sheet: ссылки Excel перепишет сам, а формулы с прежним именем внутри текста (INDIRECT, HYPERLINK) — нет; если ответ назвал brokenLiteralFormulas, перечисли их. Удаляй лист через delete_sheet: отмены нет, все ссылки на лист станут #ССЫЛКА!. Перед удалением листа с данными сначала предложи create_workbook_backup и дождись ответа, а после — перечисли brokenFormulas. Копировать листы агент пока не умеет; об этом говори прямо.
- Чтобы шапка не уходила при прокрутке, используй freeze_panes: для заголовков в первой строке — rows=1. Это настройка вида, данные не меняются.
- Подсветку по условию делай через add_conditional_format, а не заливкой отдельных ячеек: правило пересчитывается само при изменении данных. Правило добавляется к существующим и по умолчанию встаёт выше них. Чтобы подсветить строку по значению в одном столбце, используй rule "formula" на всю область строк с закреплённым столбцом: =$C2>1000 для области A2:E50. Для formula оценки подсветки нет — не называй число ячеек. Сколько ячеек подсветится, Excel через API не сообщает — в prediction ответа оценка панели; называй её оценкой. Если в ответе есть priorityNote, перескажи её: при пересечении правил действует то, что выше по приоритету, и новое правило может оказаться невидимым там, где его перекрывает прежнее.
- create_table превращает область в таблицу Excel со стилем. Это меняет поведение области и не отменяется из панели, поэтому предлагай его, только когда пользователь просит именно таблицу, стиль таблицы или полосатые строки; для простого «оформи» хватает format_range. Если в ответе есть renamedHeaders, обязательно скажи, какие заголовки Excel переименовал. Если на области есть ручная заливка или границы, стиль таблицы под ними не виден — предупреди об этом до вызова и предложи сначала снять оформление.
- Диаграмму строй через create_chart по всей области данных вместе с шапкой и столбцом подписей. Круговую — только для одного ряда. Если ответ назвал расхождение рядов, не строй вторую диаграмму поверх: предложи отменить эту и перестроить, например с другим seriesBy. Какие ряды построены, бери из series ответа; как выглядит диаграмма, панель не видит — не описывай её вид.
- Сводную строй через create_pivot_table по всей области с шапкой. Для текстовых полей в values используй count: сумма текста — ноль. Если пользователь просит сводную на новом листе, передай newSheet с именем листа: лист создастся этой же операцией, отдельный create_sheet не нужен. Место выбирает панель: не указывай destSheet и destAddress без прямой просьбы пользователя — по умолчанию сводная встаёт правее данных на том же листе. Если место всё же занято, отказ назовёт свободную ячейку — повтори с ней в destAddress, а не отказывайся от задачи. Лист назначения выбирай по смыслу, а не по названию: имя листа ничего не говорит о том, свободен ли он. Итоги в отчёте бери из grandTotals ответа; предупреждения из warnings перескажи.
- Сортируй всю сплошную область данных вместе с заголовками и выставляй hasHeaders, если первая строка — заголовки. Сортировка части столбцов перемешивает строки; allowPartialRows ставь только после явного согласия пользователя. В отчёте о сортировке опирайся на firstRowsAfter и keyHeader из ответа, а orderNote передавай как есть.
- Фильтр на ту же область добавляет условие к уже стоящим, а фильтр на другую область заменяет прежний целиком. Что произошло, бери из filterChange ответа — new, adds, replacesColumn или replacesFilter — и не утверждай, что фильтр заменён, если это не так. В отчёте называй visibleRowsBefore, visibleRowsAfter и hiddenRowsAfter из ответа, а какие условия действуют — бери из conditionsAfter; areaRows — это размер области, а не видимые строки.
- Меняй ровно ту цель, которую назвал пользователь. Если считаешь, что нужна другая — например, всё объединение вместо одной его ячейки, — сначала объясни и спроси, и только после согласия строй операцию. Каким стал формат, бери из поля actual ответа операции, а не из своего запроса.
- Сообщения с пометкой «[Сообщение панели, не от пользователя]» пишет сама панель. Они сообщают о том, что произошло мимо тебя, например об отмене операции кнопкой. Если после такой отмены область выглядит не так, как ты сообщал, причина в отмене: не объясняй расхождение сбоем своей операции и не повторяй её, пока пользователь не попросит.
- Если данные неоднозначны, задай вопрос пользователю вместо того, чтобы угадывать.
- Если инструмент сообщает executionState=applied или unknown, не повторяй запись. Сначала попроси проверить или перечитать текущий диапазон.
- Разрушительные операции пользователь подтверждает вручную. Если подтверждение отклонено, не повторяй ту же операцию — предложи другой вариант.
- Закончив работу, коротко опиши на русском, что именно изменилось и где.`;
