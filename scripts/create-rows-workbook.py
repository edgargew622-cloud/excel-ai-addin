"""Книга для проверок вставки и удаления строк: rows-smoke.xlsx.

Отдельная от остальных, потому что проверять здесь надо не целевые строки,
а то, что ломается в других местах книги. Поэтому лист «Отчёт» намеренно
ссылается на «Продажи» тремя разными способами, и каждый ведёт себя по-своему:

    B1  =СУММ(Продажи!D2:D8)      удаление внутри — диапазон молча укоротится,
                                  вставка строки 9 — новая строка не попадёт в итог
    B2  =Продажи!D4               удаление строки 4 — ссылка станет #ССЫЛКА!
    B3  =СУММ(Продажи!$D$2:$D$8)  доллары от удаления не спасают

Лист «Продажи» без таблицы Excel: таблица меняет поведение вставки и разбирается
отдельно. Строка 10 оставлена пустой намеренно — вставка ниже данных не должна
трогать ничего.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

root = Path(__file__).resolve().parent.parent
target = root / "rows-smoke.xlsx"

wb = Workbook()
sales = wb.active
sales.title = "Продажи"
sales.append(["Товар", "Категория", "Цена", "Выручка", "Доля"])
rows = [
    ("Ручка", "Канцелярия", 50, 900),
    ("Тетрадь", "Канцелярия", 80, 840),
    ("Ноутбук", "Техника", 52000, 52000),
    ("Мышь", "Техника", 900, 1800),
    ("Папка", "Канцелярия", 120, 560),
    ("Монитор", "Техника", 18000, 36000),
    ("Степлер", "Канцелярия", 200, 400),
]
for row in rows:
    sales.append(row)

# Формула внутри листа: она едет вместе со своей строкой.
for index in range(2, 9):
    sales[f"E{index}"] = f"=D{index}/SUM($D$2:$D$8)"
    sales[f"E{index}"].number_format = "0.0%"

for cell in sales[1]:
    cell.font = Font(bold=True)
sales.column_dimensions["A"].width = 14
sales.column_dimensions["B"].width = 14
sales.column_dimensions["D"].width = 12

report = wb.create_sheet("Отчёт")
report["A1"] = "Выручка всего"
report["B1"] = "=SUM(Продажи!D2:D8)"
report["A2"] = "Ноутбук отдельно"
report["B2"] = "=Продажи!D4"
report["A3"] = "Выручка всего (закреплённо)"
report["B3"] = "=SUM(Продажи!$D$2:$D$8)"
report["A4"] = "Строк в данных"
report["B4"] = "=COUNTA(Продажи!A2:A8)"
for index in range(1, 5):
    report[f"A{index}"].font = Font(bold=True)
report.column_dimensions["A"].width = 28

wb.save(target)
print(f"Книга сохранена: {target}")
