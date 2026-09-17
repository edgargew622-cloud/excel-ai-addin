"""Книга для проверок фильтра: filter-smoke.xlsx.

Отдельная от эталонной. В `reference-smoke.xlsx` данные «Продаж» оформлены
таблицей Excel, а поверх таблицы автофильтр листа не ставится; «Справочник» же
слишком мал — все товары одной категории, и по второму условию ничего
не отсеивалось, из-за чего складывание условий было плохо видно.

Здесь данные обычные, без таблицы, и подобраны так, чтобы каждое следующее
условие заметно меняло число видимых строк:

    Город = Москва      → 4 строки из 9
    + Статус = Новая    → 2 строки
    + Сумма > 500       → 1 строка

Лист «Второй блок» нужен для проверки замены: фильтр на другую область листа
заменяет прежний целиком, и это единственный случай настоящей потери условий.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

root = Path(__file__).resolve().parent.parent
target = root / "filter-smoke.xlsx"

wb = Workbook()
ws = wb.active
ws.title = "Заявки"
ws.append(["Город", "Статус", "Сумма", "Ответственный"])
rows = [
    ("Москва", "Новая", 900, "Иванов"),
    ("Москва", "Новая", 300, "Петрова"),
    ("Москва", "В работе", 1200, "Иванов"),
    ("Москва", "Закрыта", 450, "Сидоров"),
    ("Казань", "Новая", 700, "Петрова"),
    ("Казань", "В работе", 250, "Иванов"),
    ("Омск", "Новая", 1500, "Сидоров"),
    ("Омск", "Закрыта", 800, "Петрова"),
    ("Тверь", "В работе", 150, "Иванов"),
]
for row in rows:
    ws.append(row)

for cell in ws[1]:
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", fgColor="DDEBF7")
    cell.alignment = Alignment(horizontal="center")

ws.column_dimensions["A"].width = 14
ws.column_dimensions["B"].width = 14
ws.column_dimensions["D"].width = 16

# Второй блок на том же листе, отделённый пустыми строками: фильтр на него —
# другая область, и прежний фильтр будет заменён целиком.
ws["F1"] = "Товар"
ws["G1"] = "Остаток"
for cell in (ws["F1"], ws["G1"]):
    cell.font = Font(bold=True)
for index, (item, rest) in enumerate([("Кофе", 12), ("Чай", 0), ("Какао", 5)], start=2):
    ws[f"F{index}"] = item
    ws[f"G{index}"] = rest

wb.create_sheet("Пустой")
wb.calculation.fullCalcOnLoad = True
wb.save(target)
print(f"Создано: {target}")
