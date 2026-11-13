# -*- coding: utf-8 -*-
"""Генерация тестовых файлов для проверки DOC2MD (dev-скрипт, не нужен сервису)."""
import csv
import datetime as dt
import os

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "samples")
os.makedirs(OUT, exist_ok=True)

# ---------------------------------------------------------------- xlsx
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Сотрудники"
ws.append(["ФИО", "Должность", "Дата приёма", "Оклад", "Ставка", "Активен"])
ws.append(["Иванов Иван", "Инженер", dt.date(2021, 3, 15), 120000.0, 1.0, True])
ws.append(["Петрова Анна", "Аналитик", dt.date(2022, 11, 2), 95000.5, 0.5, False])
ws.append(["Сидоров Пётр", "Разработчик", dt.datetime(2023, 6, 12, 9, 30), 150000, 1.0, True])
ws.append([])
ws.append(["Итого", "", "", 365000.5, "", ""])
ws2 = wb.create_sheet("Итоги")
ws2.append(["Квартал", "Выручка"])
ws2.append(["Q1 2025", 1234500])
ws2.append(["Q2 2025", 2345600])
ws2.merge_cells("A4:B4")
ws2["A4"] = "Сводка по году: рост 90%"
wb.save(os.path.join(OUT, "sample.xlsx"))
print("sample.xlsx ok")

# ---------------------------------------------------------------- csv
p = os.path.join(OUT, "sample.csv")
with open(p, "w", encoding="utf-8", newline="") as f:
    w = csv.writer(f, delimiter=";")
    w.writerow(["id", "Товар", "Количество", "Цена"])
    w.writerow([1, "Кофе", 3, "450,50"])
    w.writerow([2, "Чай", 10, "120"])
    w.writerow([3, "Сахар", 1, "89,90"])
print("sample.csv ok")

# ---------------------------------------------------------------- docx
import docx  # noqa: E402  (dev-зависимость)

d = docx.Document()
d.add_heading("Отчёт о проекте Альфа", level=1)
d.add_heading("Резюме", level=2)
d.add_paragraph("Проект завершён в срок. ")
p = d.add_paragraph("Ключевой результат: ")
p.add_run("сокращение времени обработки").bold = True
p.add_run(" на 40%.")
d.add_heading("Задачи", level=2)
for t in ["Сбор требований", "Разработка прототипа", "Опытная эксплуатация"]:
    d.add_paragraph(t, style="List Bullet")
d.add_heading("Дальнейшие шаги", level=2)
for t in ["Подписать акт", "Передать документацию"]:
    d.add_paragraph(t, style="List Number")
d.add_heading("Бюджет", level=2)
tbl = d.add_table(rows=3, cols=3)
tbl.style = "Table Grid"
data = [["Статья", "План", "Факт"], ["Разработка", "1 000 000", "980 000"], ["Тестирование", "300 000", "310 000"]]
for r, row in enumerate(data):
    for c, val in enumerate(row):
        tbl.cell(r, c).text = val
d.save(os.path.join(OUT, "sample.docx"))
print("sample.docx ok")

# ---------------------------------------------------------------- pdf
import fitz  # noqa: E402

FONT = r"C:\Windows\Fonts\arial.ttf"
doc = fitz.open()
FOOTER = "Внутренний отчёт — конфиденциально"

PAGE_TEXT = {
    1: ["Раздел 1. Постановка задачи: автоматизация отчётности за квартал.",
        "Источник данных — учётная система, выгрузка выполняется еженедельно."],
    2: ["Раздел 2. Ход выполнения: все этапы пройдены, отклонений по срокам нет.",
        "Команда подтверждает готовность к опытной эксплуатации."],
    3: ["Раздел 3. Итоги: цель проекта достигнута, метрики улучшены.",
        "Рекомендация — перейти к промышленной эксплуатации с 1 декабря."],
}

for i in range(1, 4):
    page = doc.new_page()  # A4 595x842
    y = 72
    page.insert_text((72, y), "Отчёт о проекте Альфа", fontname="F0", fontfile=FONT, fontsize=18)
    y += 30
    page.insert_text((72, y), f"Страница {i} из 3", fontname="F0", fontfile=FONT, fontsize=9)
    y += 24
    for para in PAGE_TEXT[i]:
        page.insert_text((72, y), para, fontname="F0", fontfile=FONT, fontsize=11)
        y += 20
    # «таблица»: сетка линий + текст
    x0, y0, x1 = 72, y + 14, 400
    rowh, colw = 20, 110
    rows_data = [["Статья", "План", "Факт"],
                 ["Разработка", "1000", "980"],
                 ["Тестирование", "300", "310"]]
    for r in range(len(rows_data) + 1):
        yy = y0 + r * rowh
        page.draw_line((x0, yy), (x1, yy))
    for c in range(4):
        xx = x0 + c * colw
        page.draw_line((xx, y0), (xx, y0 + len(rows_data) * rowh))
    for r, row in enumerate(rows_data):
        for c, val in enumerate(row):
            page.insert_text((x0 + c * colw + 4, y0 + r * rowh + 15),
                             val, fontname="F0", fontfile=FONT, fontsize=10)
    y = y0 + len(rows_data) * rowh + 40
    page.insert_text((72, y), PAGE_TEXT[i][1], fontname="F0", fontfile=FONT, fontsize=11)
    # колонтитул и номер страницы
    page.insert_text((72, 810), FOOTER, fontname="F0", fontfile=FONT, fontsize=8)
    page.insert_text((285, 825), str(i), fontname="F0", fontfile=FONT, fontsize=9)

doc.save(os.path.join(OUT, "sample.pdf"))
doc.close()
print("sample.pdf ok")
print("Готово:", OUT)
