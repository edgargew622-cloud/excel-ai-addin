import { test } from "node:test";
import assert from "node:assert/strict";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { Markdown, actionsWord, toolLabel } from "./chatView";

const html = (text: string) => renderToStaticMarkup(createElement(Markdown, { text }));

test("markdown: bold, code, lists and tables become elements, not raw asterisks", () => {
  // Ответ Claude из беседы «Книга16» 02.10.2026.
  const out = html("Готово.\n\n| | Янв | Фев |\n|---|---|---|\n| **Итого** | 420 | 455 |\n\n- **Формулы:** `=C5/B5-1`\n- Январь: прочерк");
  assert.match(out, /<table>/);
  assert.match(out, /<td><strong>Итого<\/strong><\/td>/);
  assert.match(out, /<ul><li><strong>Формулы:<\/strong> <code>=C5\/B5-1<\/code><\/li>/);
  assert.doesNotMatch(out, /\*\*/);
  assert.doesNotMatch(out, /\|---/);
});

test("markdown: html from the model or a cell is shown as text, never as markup", () => {
  const out = html("Ячейка A1: <img src=x onerror=alert(1)> и [ссылка](javascript:alert(1))");
  assert.doesNotMatch(out, /<img/);
  assert.match(out, /&lt;img/);
  assert.doesNotMatch(out, /href="javascript/);
});

test("markdown: underscores in names stay as they are", () => {
  assert.match(html("Лист Итого_2024 и столбец Цена_руб"), /Итого_2024 и столбец Цена_руб/);
});

test("actions: Russian labels and counts", () => {
  assert.equal(toolLabel("format_range"), "Оформление");
  assert.equal(toolLabel("some_new_tool"), "some_new_tool");
  assert.equal(actionsWord(1), "1 действие");
  assert.equal(actionsWord(3), "3 действия");
  assert.equal(actionsWord(5), "5 действий");
  assert.equal(actionsWord(12), "12 действий");
});
