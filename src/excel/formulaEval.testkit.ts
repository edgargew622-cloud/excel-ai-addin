/**
 * Вычислитель формул для тестов моделей: + − * / ^, сравнения, ссылки,
 * SUM по диапазону, MAX, ROUND, IF и текстовые "". Этого хватает, чтобы
 * проверить, что формулы на листе дают ровно расчёт панели, — без Excel.
 */

import { columnIndex, columnLetters } from "./formulaFill";

type Value = number | string | boolean;

export function evaluateGrid(rows: readonly (readonly { formula: string | number }[])[]): Value[][] {
  const cache = new Map<string, Value>();
  const cellValue = (address: string): Value => {
    const key = address.replace(/\$/g, "");
    if (cache.has(key)) return cache.get(key)!;
    const match = /^([A-Z]+)(\d+)$/.exec(key)!;
    const formula = rows[Number(match[2]) - 1]?.[columnIndex(match[1]) - 1]?.formula ?? "";
    const value = typeof formula === "string" && formula.startsWith("=") ? calc(formula.slice(1)) : formula;
    cache.set(key, value);
    return value;
  };
  const num = (value: Value) => (typeof value === "number" ? value : typeof value === "boolean" ? Number(value) : 0);

  function calc(text: string): Value {
    let i = 0;
    const rest = () => text.slice(i);
    const compare = (): Value => {
      const left = additive();
      const op = /^(<=|>=|<>|<|>|=)/.exec(rest());
      if (!op) return left;
      i += op[0].length;
      const right = additive();
      const a = num(left), b = num(right);
      return op[0] === "<=" ? a <= b : op[0] === ">=" ? a >= b : op[0] === "<" ? a < b : op[0] === ">" ? a > b : op[0] === "=" ? a === b : a !== b;
    };
    const additive = (): Value => {
      let value = term();
      while (text[i] === "+" || text[i] === "-") value = text[i++] === "+" ? num(value) + num(term()) : num(value) - num(term());
      return value;
    };
    const term = (): Value => {
      let value = power();
      while (text[i] === "*" || text[i] === "/") value = text[i++] === "*" ? num(value) * num(power()) : num(value) / num(power());
      return value;
    };
    const power = (): Value => {
      let value = unary();
      while (text[i] === "^") { i++; value = num(value) ** num(unary()); }
      return value;
    };
    const unary = (): Value => {
      if (text[i] === "-") { i++; return -num(unary()); }
      return atom();
    };
    const args = (): Value[] => {
      const list: Value[] = [];
      while (text[i] !== ")") { list.push(compare()); if (text[i] === ",") i++; }
      i++;
      return list;
    };
    const atom = (): Value => {
      if (text[i] === "(") { i++; const value = compare(); i++; return value; }
      if (text[i] === '"') { const end = text.indexOf('"', i + 1); const value = text.slice(i + 1, end); i = end + 1; return value; }
      const fn = /^(MAX|ROUND|SUM|IF)\(/.exec(rest());
      if (fn) {
        i += fn[0].length;
        if (fn[1] === "SUM") {
          const range = /^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\)/.exec(rest())!;
          i += range[0].length;
          let total = 0;
          for (let c = columnIndex(range[1]); c <= columnIndex(range[3]); c++) {
            for (let r = Number(range[2]); r <= Number(range[4]); r++) total += num(cellValue(`${columnLetters(c)}${r}`));
          }
          return total;
        }
        if (fn[1] === "IF") {
          const [condition, yes, no] = args();
          return condition ? yes : no;
        }
        const [first, second] = args().map(num);
        return fn[1] === "MAX" ? Math.max(first, second) : Math.round(first * 10 ** second) / 10 ** second;
      }
      const ref = /^\$?[A-Z]+\$?\d+/.exec(rest());
      if (ref) { i += ref[0].length; return cellValue(ref[0]); }
      const number = /^\d+(\.\d+)?/.exec(rest())!;
      i += number[0].length;
      return Number(number[0]);
    };
    return compare();
  }
  return rows.map((row, r) => row.map((_, c) => cellValue(`${columnLetters(c + 1)}${r + 1}`)));
}
