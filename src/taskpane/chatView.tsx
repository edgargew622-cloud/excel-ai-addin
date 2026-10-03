/**
 * Вид чата (03.10.2026): разметка ответов модели, понятные названия
 * действий и значки. Ответы приходят в Markdown — прежде панель показывала
 * их как есть, со звёздочками и таблицами из вертикальных черт.
 *
 * Разбор свой и только в элементы React: HTML из ответа модели в панель
 * не попадает (в ответ может попасть текст из ячеек книги).
 */

import { Fragment, type ReactNode } from "react";

/* ------------------------------------------------------------ Markdown */

function inline(text: string, key = 0): ReactNode[] {
  const out: ReactNode[] = [];
  // Код, жирный, курсив звёздочкой, ссылка. Подчёркивания не трогаем:
  // в ответах много имён вроде «Итого_2024».
  const pattern = /(`[^`]+`)|(\*\*[^*]+\*\*)|(\*[^*\s][^*]*\*)|(\[[^\]]+\]\((https?:\/\/[^)\s]+)\))/g;
  let last = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(text))) {
    if (match.index > last) out.push(text.slice(last, match.index));
    const token = match[0];
    const k = `${key}-${match.index}`;
    if (match[1]) out.push(<code key={k}>{token.slice(1, -1)}</code>);
    else if (match[2]) out.push(<strong key={k}>{inline(token.slice(2, -2), match.index)}</strong>);
    else if (match[3]) out.push(<em key={k}>{inline(token.slice(1, -1), match.index)}</em>);
    else {
      const label = token.slice(1, token.indexOf("]("));
      out.push(<a key={k} href={match[5]} target="_blank" rel="noreferrer">{label}</a>);
    }
    last = match.index + token.length;
  }
  if (last < text.length) out.push(text.slice(last));
  return out;
}

const cells = (line: string) => line.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map((cell) => cell.trim());
const isTableRule = (line: string) => /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(line);

export function Markdown({ text }: { text: string }) {
  const lines = text.replace(/\r\n/g, "\n").split("\n");
  const blocks: ReactNode[] = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) { i++; continue; }

    if (/^\s*```/.test(line)) {
      const body: string[] = [];
      i++;
      while (i < lines.length && !/^\s*```/.test(lines[i])) body.push(lines[i++]);
      i++;
      blocks.push(<pre key={i} className="md-code">{body.join("\n")}</pre>);
      continue;
    }

    if (line.includes("|") && i + 1 < lines.length && isTableRule(lines[i + 1])) {
      const head = cells(line);
      const rows: string[][] = [];
      i += 2;
      while (i < lines.length && lines[i].includes("|") && lines[i].trim()) rows.push(cells(lines[i++]));
      blocks.push(
        <div key={i} className="md-table">
          <table>
            <thead><tr>{head.map((cell, c) => <th key={c}>{inline(cell)}</th>)}</tr></thead>
            <tbody>{rows.map((row, r) => <tr key={r}>{row.map((cell, c) => <td key={c}>{inline(cell)}</td>)}</tr>)}</tbody>
          </table>
        </div>
      );
      continue;
    }

    const heading = /^(#{1,4})\s+(.*)$/.exec(line);
    if (heading) {
      blocks.push(<div key={i} className={`md-h md-h${heading[1].length}`}>{inline(heading[2])}</div>);
      i++;
      continue;
    }

    if (/^\s*([-*•]|\d+[.)])\s+/.test(line)) {
      const ordered = /^\s*\d+[.)]/.test(line);
      const items: { text: string; nested: boolean }[] = [];
      while (i < lines.length && /^\s*([-*•]|\d+[.)])\s+/.test(lines[i])) {
        const nested = /^\s{2,}/.test(lines[i]);
        items.push({ text: lines[i].replace(/^\s*([-*•]|\d+[.)])\s+/, ""), nested });
        i++;
        // Продолжение пункта на следующей строке с отступом.
        while (i < lines.length && /^\s{2,}\S/.test(lines[i]) && !/^\s*([-*•]|\d+[.)])\s+/.test(lines[i])) {
          items[items.length - 1].text += " " + lines[i].trim();
          i++;
        }
      }
      const List = ordered ? "ol" : "ul";
      blocks.push(
        <List key={i}>
          {items.map((item, n) => <li key={n} className={item.nested ? "nested" : undefined}>{inline(item.text)}</li>)}
        </List>
      );
      continue;
    }

    const para: string[] = [];
    while (
      i < lines.length && lines[i].trim() &&
      !/^\s*```/.test(lines[i]) && !/^#{1,4}\s/.test(lines[i]) && !/^\s*([-*•]|\d+[.)])\s+/.test(lines[i]) &&
      !(lines[i].includes("|") && i + 1 < lines.length && isTableRule(lines[i + 1]))
    ) para.push(lines[i++]);
    blocks.push(
      <p key={i}>
        {para.map((part, n) => <Fragment key={n}>{n > 0 && <br />}{inline(part, n)}</Fragment>)}
      </p>
    );
  }
  return <>{blocks}</>;
}

/* ------------------------------------------------------- действия агента */

const TOOL_LABELS: Record<string, string> = {
  get_range_values: "Чтение ячеек",
  get_range_details: "Чтение подробностей",
  get_sheet_overview: "Обзор листа",
  list_sheets: "Список листов",
  get_active_context: "Текущее выделение",
  search_workbook: "Поиск по книге",
  profile_range: "Разбор данных",
  analyze_range: "Анализ данных",
  audit_workbook: "Проверка книги",
  get_conditional_formats: "Чтение правил оформления",
  measure_workbook_export: "Оценка размера",
  recall_snapshot: "Сверка с прежним",
  list_files: "Список файлов",
  read_file: "Чтение файла",
  web_search: "Поиск в интернете",
  read_web_page: "Чтение страницы",
  get_scenario: "Сценарий",
  set_range_values: "Запись в ячейки",
  set_ranges_values: "Запись в ячейки",
  fill_range: "Заполнение",
  format_range: "Оформление",
  sort_range: "Сортировка",
  apply_filter: "Фильтр",
  insert_rows: "Вставка строк",
  delete_rows: "Удаление строк",
  insert_columns: "Вставка столбцов",
  delete_columns: "Удаление столбцов",
  freeze_panes: "Закрепление областей",
  add_conditional_format: "Условное форматирование",
  move_conditional_format: "Перенос условного форматирования",
  create_table: "Создание таблицы",
  convert_table_to_range: "Таблица в обычный диапазон",
  create_chart: "Диаграмма",
  format_chart: "Оформление диаграммы",
  create_pivot_table: "Сводная таблица",
  create_sheet: "Новый лист",
  copy_sheet: "Копия листа",
  rename_sheet: "Переименование листа",
  delete_sheet: "Удаление листа",
  trim_text: "Лишние пробелы",
  convert_values: "Преобразование значений",
  change_case: "Регистр текста",
  remove_duplicates: "Удаление дублей",
  set_page_layout: "Параметры печати",
  group_rows_columns: "Группировка",
  set_data_validation: "Проверка ввода",
  apply_color_convention: "Цвета модели",
  add_multiples: "Мультипликаторы",
  add_share_growth: "Доли и рост",
  add_comparison: "Сравнение",
  build_three_statement_model: "Модель трёх отчётов",
  build_dcf_model: "Модель DCF",
  build_lbo_model: "Модель LBO",
  import_file_layout: "Копия файла на лист",
  import_file_table: "Таблица из файла",
  create_workbook_backup: "Резервная копия",
  remember_preference: "Запомнить",
  save_scenario: "Сохранение сценария",
  __read_sheets: "Чтение других листов"
};

export const toolLabel = (name: string) => TOOL_LABELS[name] ?? name;

/** «4 действия», «1 действие». */
export function actionsWord(n: number): string {
  const word = n % 10 === 1 && n % 100 !== 11 ? "действие" : [2, 3, 4].includes(n % 10) && ![12, 13, 14].includes(n % 100) ? "действия" : "действий";
  return `${n} ${word}`;
}

/* ---------------------------------------------------------------- значки */

const svg = (path: ReactNode, size = 16) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{path}</svg>
);

export const Icon = {
  send: () => svg(<><path d="M12 19V5" /><path d="m5 12 7-7 7 7" /></>),
  stop: () => svg(<rect x="7" y="7" width="10" height="10" rx="1.5" fill="currentColor" stroke="none" />),
  clip: () => svg(<path d="m21 11-8.5 8.5a5 5 0 0 1-7-7L14 4a3.3 3.3 0 0 1 4.7 4.7l-8.5 8.5a1.7 1.7 0 0 1-2.4-2.4L15.5 7" />),
  undo: () => svg(<><path d="M9 14 4 9l5-5" /><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11" /></>, 14),
  check: () => svg(<path d="M20 6 9 17l-5-5" />, 13),
  cross: () => svg(<><path d="M18 6 6 18" /><path d="m6 6 12 12" /></>, 13),
  alert: () => svg(<><path d="M12 9v4" /><path d="M12 17h.01" /><circle cx="12" cy="12" r="9" /></>, 13),
  chevron: () => svg(<path d="m9 18 6-6-6-6" />, 13),
  sheet: () => svg(<><rect x="3" y="3" width="18" height="18" rx="2" /><path d="M3 9h18M9 21V9" /></>, 13),
  coin: () => svg(<><circle cx="12" cy="12" r="9" /><path d="M14.5 9.5c-.5-1-1.5-1.5-2.5-1.5-1.5 0-2.5.8-2.5 2s1 1.7 2.5 2 2.5.8 2.5 2-1 2-2.5 2c-1 0-2-.5-2.5-1.5M12 6.5v1.5M12 16v1.5" /></>, 13),
  plus: () => svg(<><path d="M12 5v14" /><path d="M5 12h14" /></>, 14),
  shield: () => svg(<path d="M12 3 5 6v5c0 4.5 3 8 7 10 4-2 7-5.5 7-10V6z" />, 15)
};

/** Знак am.AI — тот же файл, что на кнопке в ленте Excel (public/assets). */
export function Logo({ size = 24, className = "" }: { size?: number; className?: string }) {
  return (
    <img
      className={`logo ${className}`.trim()}
      src={size > 64 ? "assets/amai-128.png" : "assets/amai-64.png"}
      width={size}
      height={size}
      alt=""
      aria-hidden="true"
    />
  );
}

/** Надпись «am.AI»: «AI» — градиентом знака, как на логотипе. */
export function Wordmark() {
  return <span className="wordmark">am.<b>AI</b></span>;
}

/** Звезда со знака — индикатор «думает». */
export function Sparkle({ size = 16 }: { size?: number }) {
  return (
    <svg className="sparkle" width={size} height={size} viewBox="0 0 24 24" aria-hidden="true">
      <defs>
        <linearGradient id="amai-spark" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stopColor="#3ddc84" />
          <stop offset="0.55" stopColor="#14b8a6" />
          <stop offset="1" stopColor="#06b6d4" />
        </linearGradient>
      </defs>
      <path d="M12 1.5c.6 4.8 2.7 7.5 9.5 10.5-6.8 3-8.9 5.7-9.5 10.5-.6-4.8-2.7-7.5-9.5-10.5 6.8-3 8.9-5.7 9.5-10.5z" fill="url(#amai-spark)" />
    </svg>
  );
}
