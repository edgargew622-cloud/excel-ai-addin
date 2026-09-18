"""Книга для проверок диаграмм: chart-smoke.xlsx.

«Продажи» — привычная таблица: месяцы в первом столбце, три числовых ряда.
На ней видно, правильно ли Excel понял шапку и подписи, а круговая по ней
обязана предупредить, что нарисует только первый ряд.

«Доли» — одна колонка чисел для честной круговой.

«Широкая» — месяцы идут по столбцам, а не по строкам: без seriesBy=rows
получится двенадцать рядов по одной точке, и это должно быть названо.

«С комментарием» — рядом с числами текстовый столбец: если захватить его
в область, Excel построит ряд из нулей.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

root = Path(__file__).resolve().parent.parent
target = root / "chart-smoke.xlsx"

wb = Workbook()
sales = wb.active
sales.title = "Продажи"
sales.append(["Месяц", "Выручка", "Расходы", "Прибыль"])
months = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь"]
revenue = [120, 150, 170, 160, 190, 210]
costs = [80, 90, 95, 100, 110, 120]
for month, r, c in zip(months, revenue, costs):
    sales.append([month, r, c, r - c])
for cell in sales[1]:
    cell.font = Font(bold=True)

shares = wb.create_sheet("Доли")
shares.append(["Канал", "Выручка"])
for row in [("Магазин", 540), ("Сайт", 310), ("Маркетплейс", 260), ("Опт", 90)]:
    shares.append(row)
for cell in shares[1]:
    cell.font = Font(bold=True)

wide = wb.create_sheet("Широкая")
wide.append(["Показатель"] + months)
wide.append(["Выручка"] + revenue)
for cell in wide[1]:
    cell.font = Font(bold=True)

noted = wb.create_sheet("С комментарием")
noted.append(["Месяц", "Выручка", "Комментарий"])
for month, r, note in zip(months, revenue, ["план", "план", "факт", "факт", "прогноз", "прогноз"]):
    noted.append([month, r, note])
for cell in noted[1]:
    cell.font = Font(bold=True)

wb.save(target)
print(f"Книга сохранена: {target}")
