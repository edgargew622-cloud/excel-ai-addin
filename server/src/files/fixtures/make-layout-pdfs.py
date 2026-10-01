# Проверочные PDF для переноса «как есть»: сетка, бланк, без линий, две страницы.
import json, sys
from fpdf import FPDF

F = "C:/Windows/Fonts/"
out = sys.argv[1]

def doc():
    p = FPDF(unit="pt", format="A4")
    p.add_font("A", "", F + "arial.ttf")
    p.add_font("A", "B", F + "arialbd.ttf")
    p.set_auto_page_break(False)
    p.add_page()
    return p

expect = {}

# 1. Таблица с сеткой.
p = doc()
p.set_font("A", "B", 14); p.set_xy(40, 40); p.cell(515, 20, "Реестр платежей за сентябрь 2026", align="C")
p.set_font("A", "", 9); p.set_xy(40, 62); p.cell(515, 12, "Организация: ООО «Пример», ИНН 7700000001")
rows = [["№", "Дата", "Контрагент", "Назначение платежа", "Сумма, руб."],
        ["1", "01.09.2026", "ООО «Альфа»", "Оплата по счёту 15", "12 500,00"],
        ["2", "05.09.2026", "ИП Иванов", "Аренда офиса", "45 000,00"],
        ["3", "12.09.2026", "АО «Бета»", "Поставка бумаги", "3 210,50"],
        ["4", "28.09.2026", "ООО «Гамма»", "Услуги связи", "1 899,99"]]
widths = [30, 75, 120, 200, 90]
y = 90
for r, row in enumerate(rows):
    x = 40
    for c, text in enumerate(row):
        p.set_font("A", "B" if r == 0 else "", 10)
        p.set_xy(x, y); p.cell(widths[c], 20, text, border=1, align="R" if (c == 4 and r > 0) else ("C" if r == 0 else "L"))
        x += widths[c]
    y += 20
p.set_font("A", "B", 10); p.set_xy(40, y); p.cell(425, 20, "Итого", border=1, align="R"); p.set_xy(465, y); p.cell(90, 20, "62 610,49", border=1, align="R")
p.output(out + "/grid.pdf")
expect["grid"] = {"values": rows + [["Итого", "62 610,49"]]}

# 2. Бланк: объединения, многострочная ячейка, пустые поля.
p = doc()
p.set_font("A", "B", 12); p.set_xy(40, 40); p.cell(515, 18, "ПУТЕВОЙ ЛИСТ № 0042", align="C")
p.set_font("A", "", 9); p.set_xy(40, 58); p.cell(515, 12, "с 30.09.2026 по 01.10.2026", align="C")
def box(x, y, w, h, text="", bold=False, align="L"):
    p.rect(x, y, w, h)
    if text:
        p.set_font("A", "B" if bold else "", 9); p.set_xy(x + 2, y + 2); p.multi_cell(w - 4, 11, text, align=align)
top = 80
box(40, top, 515, 18, "Сведения о транспортном средстве", True, "C")
box(40, top + 18, 150, 18, "Марка"); box(190, top + 18, 365, 18, "SOLARIS")
box(40, top + 36, 150, 18, "Госномер"); box(190, top + 36, 120, 18, "Е000ЕЕ000"); box(310, top + 36, 100, 18, "Гаражный №"); box(410, top + 36, 145, 18, "17")
box(40, top + 54, 515, 18, "Водитель", True, "C")
box(40, top + 72, 150, 36, "ФИО"); box(190, top + 72, 365, 36, "Иванов Иван Иванович,\nудостоверение 99 00 000000")
box(40, top + 108, 257, 18, "Выезд", True, "C"); box(297, top + 108, 258, 18, "Возвращение", True, "C")
box(40, top + 126, 90, 18, "Время"); box(130, top + 126, 167, 18, "08:00"); box(297, top + 126, 90, 18, "Время"); box(387, top + 126, 168, 18, "")
box(40, top + 144, 90, 18, "Одометр, км"); box(130, top + 144, 167, 18, "114 106"); box(297, top + 144, 90, 18, "Одометр, км"); box(387, top + 144, 168, 18, "")
p.set_font("A", "", 8); p.set_xy(40, top + 170); p.cell(515, 10, "Особые отметки: —")
p.output(out + "/form.pdf")

# 3. Таблица без линий, числа по правому краю.
p = doc()
p.set_font("A", "B", 12); p.set_xy(40, 40); p.cell(300, 16, "Остатки на складе")
cols = [(40, 160, "L"), (200, 60, "R"), (260, 50, "L"), (330, 80, "R"), (430, 90, "R")]
data = [["Товар", "Кол-во", "Ед.", "Цена", "Сумма"],
        ["Бумага A4", "120", "пач.", "310,00", "37 200,00"],
        ["Ручка шариковая", "1 500", "шт.", "12,50", "18 750,00"],
        ["Скрепки", "8", "уп.", "45,00", "360,00"],
        ["Папка-регистратор", "35", "шт.", "189,90", "6 646,50"]]
y = 70
for r, row in enumerate(data):
    for (x, w, a), text in zip(cols, row):
        p.set_font("A", "B" if r == 0 else "", 10); p.set_xy(x, y); p.cell(w, 14, text, align=a)
    y += 16
p.output(out + "/plain.pdf")

# 4. Две страницы с повтором шапки.
p = doc()
head = ["Дата", "Операция", "Приход", "Расход", "Остаток"]
w2 = [70, 200, 80, 80, 85]
def header(y):
    x = 40
    for c, t in enumerate(head):
        p.set_font("A", "B", 9); p.set_xy(x, y); p.cell(w2[c], 16, t, border=1, align="C"); x += w2[c]
y = 50; header(y); y += 16; balance = 1000.0
for i in range(70):
    if y > 780:
        p.add_page(); y = 50; header(y); y += 16
    inc = (i % 3 == 0) * (100 + i); out_ = (i % 3 != 0) * (20 + i); balance += inc - out_
    vals = [f"{(i % 28) + 1:02d}.10.2026", f"Операция {i + 1}", f"{inc:,.2f}".replace(",", " ").replace(".", ",") if inc else "", f"{out_:,.2f}".replace(",", " ").replace(".", ",") if out_ else "", f"{balance:,.2f}".replace(",", " ").replace(".", ",")]
    x = 40
    for c, t in enumerate(vals):
        p.set_font("A", "", 9); p.set_xy(x, y); p.cell(w2[c], 16, t, border=1, align="R" if c >= 2 else "L"); x += w2[c]
    y += 16
p.output(out + "/multipage.pdf")
print("ok")
