/**
 * Буквы столбцов Excel.
 *
 * Раньше здесь же был собственный сдвиг формул — запасной путь заполнения
 * на случай, когда протяжка Excel отказывала. Он принимал похожие на адрес
 * имена за ссылки: `=Q1!A1` становилось `=Q2!A2`, а `=SUM(Sales[Q1])` —
 * `=SUM(Sales[Q2])` (план стабилизации, S1). Заполнение теперь делает сам
 * Excel через формулы в виде R1C1, и разбор формул здесь больше не нужен.
 */

export function columnLetters(index: number): string {
  let value = "";
  let left = index;
  while (left > 0) {
    const remainder = (left - 1) % 26;
    value = String.fromCharCode(65 + remainder) + value;
    left = Math.floor((left - 1) / 26);
  }
  return value;
}

export function columnIndex(letters: string): number {
  let index = 0;
  for (const letter of letters.toUpperCase()) index = index * 26 + (letter.charCodeAt(0) - 64);
  return index;
}
