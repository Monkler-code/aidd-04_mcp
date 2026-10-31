#!/usr/bin/env python3
"""Генерация отчёта о задачах в .xlsx: два листа — «Все задачи» и «План на неделю»."""
import re
import zipfile
from datetime import date, datetime, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

TASKS = [
    (1, "Позвонить по поводу договор", "low", None, False),
    (2, "Протестировать билеты на поезд", "low", None, True),
    (3, "Протестировать молоко и хлеб", "low", "2026-11-22", False),
    (4, "Позвонить по поводу анкету", "normal", None, False),
    (5, "Закрыть молоко и хлеб", "normal", None, True),
    (6, "Спланировать доклад для лекции", "normal", "2026-11-13", True),
    (7, "Купить отчёт за квартал", "normal", "2026-11-10", True),
    (8, "Отправить договор", "low", None, True),
    (9, "Забронировать презентацию", "high", "2026-11-11", True),
    (10, "Обновить план на неделю", "low", "2026-10-23", False),
    (11, "Купить билеты на поезд", "high", "2026-11-20", True),
    (12, "Записаться на договор", "high", "2026-10-23", False),
    (13, "Купить резервную копию", "high", "2026-10-17", True),
    (14, "Разобрать обновления системы", "normal", "2026-11-05", False),
    (15, "Закрыть договор", "low", "2026-10-16", False),
    (16, "Согласовать билеты на поезд", "high", "2026-10-23", True),
    (17, "Спланировать документацию", "high", "2026-11-04", False),
    (18, "Обновить встречу с командой", "normal", "2026-11-16", True),
    (19, "Закрыть отчёт за квартал", "normal", "2026-10-17", False),
    (20, "Купить резервную копию", "high", "2026-11-06", True),
    (21, "Пропылесосить полы", "low", None, True),
    (22, "Записаться на документацию", "low", "2026-10-22", True),
    (23, "Распечатать презентацию", "high", None, False),
    (24, "Согласовать документацию", "high", None, False),
    (25, "Заказать документацию", "normal", "2026-11-24", False),
    (26, "Позвонить по поводу обновления системы", "normal", "2026-10-21", True),
    (27, "Обновить презентацию", "normal", "2026-11-11", True),
    (28, "Записаться на договор", "high", "2026-11-28", False),
    (29, "Починить презентацию", "high", None, True),
    (30, "Протестировать резервную копию", "high", "2026-11-13", True),
    (31, "Проверить документацию", "low", "2026-11-08", False),
    (32, "Позвонить по поводу подарок коллеге", "normal", "2026-11-13", True),
    (33, "Доработать визит к врачу", "high", "2026-10-25", False),
    (34, "Починить билеты на поезд", "high", None, False),
    (35, "Отправить документацию", "low", "2026-11-10", False),
    (36, "Согласовать посылку", "high", "2026-10-25", True),
    (37, "Прочитать молоко и хлеб", "normal", "2026-11-15", True),
    (38, "Записаться на посылку", "normal", "2026-11-24", False),
    (39, "Купить доклад для лекции", "normal", "2026-11-04", True),
    (40, "Обновить молоко и хлеб", "high", None, True),
    (41, "Проверить заявку на отпуск", "normal", "2026-11-28", False),
    (42, "Распечатать отчёт за квартал", "high", "2026-11-06", False),
    (43, "Записаться на план на неделю", "low", "2026-11-24", True),
    (44, "Подготовить билеты на поезд", "high", None, False),
    (45, "Разобрать доклад для лекции", "high", "2026-10-27", True),
    (46, "Распечатать доклад для лекции", "low", "2026-10-31", False),
    (47, "Забронировать ревью кода", "low", "2026-10-28", True),
    (48, "Проверить встречу с командой", "low", "2026-10-26", False),
    (49, "Купить встречу с командой", "high", "2026-10-29", True),
    (50, "Заказать резервную копию", "high", "2026-11-28", False),
    (51, "Починить подарок коллеге", "normal", "2026-11-09", False),
    (52, "Оплатить презентацию", "normal", "2026-11-05", True),
    (53, "Оплатить отчёт за квартал", "low", "2026-11-27", False),
    (54, "Проверить договор", "normal", None, True),
    (55, "Согласовать анкету", "low", "2026-11-21", False),
    (56, "Спланировать документацию", "normal", "2026-11-26", False),
    (57, "Обновить обновления системы", "normal", None, False),
    (58, "Обновить договор", "normal", "2026-10-24", False),
    (59, "Оплатить резервную копию", "low", "2026-10-27", False),
    (60, "Позвонить по поводу билеты на поезд", "normal", None, True),
    (61, "Распечатать счёт за интернет", "low", "2026-11-18", False),
    (62, "Проверить документацию", "low", "2026-10-21", False),
    (63, "Отправить визит к врачу", "normal", "2026-11-15", True),
]

PRIORITIES = {"high": "Высокий", "normal": "Средний", "low": "Низкий"}
PRIORITY_ORDER = {"high": 0, "normal": 1, "low": 2}
STATUS = {True: "Выполнено", False: "В работе"}

TODAY = date(2026, 10, 31)                            # суббота
WEEK_START = TODAY - timedelta(days=TODAY.weekday())  # понедельник текущей недели
WEEK_END = WEEK_START + timedelta(days=6)             # воскресенье текущей недели

FILE_PATH = "/home/monkler/aidd_week_04_mcp/practice/agent/.qwen/skills/report.xlsx"

# Фиксированное время: без него openpyxl пишет текущие даты в core.xml и метки
# времени в ZIP, из-за чего перезапуск скрипта даёт другой байтовый результат.
FIXED_TIME = datetime(2016, 1, 1, 0, 0, 0)

thin = Side(style="thin", color="B0B0B0")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
header_fill = PatternFill("solid", fgColor="1F4E78")
header_font = Font(bold=True, color="FFFFFF", size=11)
high_fill = PatternFill("solid", fgColor="F8CBAD")
normal_fill = PatternFill("solid", fgColor="FCE4D6")
low_fill = PatternFill("solid", fgColor="C6EFCE")
alt_fill = PatternFill("solid", fgColor="D9EAF7")


def fmt(d):
    if isinstance(d, str):
        d = date.fromisoformat(d)
    return d.strftime("%d.%m.%Y") if d else "—"


def priority_fill(priority):
    return {"Высокий": high_fill, "Средний": normal_fill, "Низкий": low_fill}.get(priority)


def save_workbook(wb, path):
    """Сохраняет книгу в xlsx и переписывает ZIP с фиксированными метками времени,
    чтобы перезапуск скрипта давал побайтово тот же результат."""
    wb.save(path)
    fixed_date = b"2016-01-01T00:00:00Z"
    data = {}
    with zipfile.ZipFile(path) as zin:
        for info in zin.infolist():
            blob = zin.read(info.filename)
            if info.filename == "docProps/core.xml":
                blob = re.sub(rb"<dcterms:created[^>]*>[^<]+</dcterms:created>",
                              b'<dcterms:created xsi:type="dcterms:W3CDTF">' + fixed_date + b"</dcterms:created>", blob)
                blob = re.sub(rb"<dcterms:modified[^>]*>[^<]+</dcterms:modified>",
                              b'<dcterms:modified xsi:type="dcterms:W3CDTF">' + fixed_date + b"</dcterms:modified>", blob)
            data[info.filename] = blob
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, blob in data.items():
            zi = zipfile.ZipInfo(name, date_time=(2016, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o644 << 16
            zout.writestr(zi, blob)


def build_report(path=None):
    """Строит и сохраняет отчёт в xlsx (по умолчанию FILE_PATH). Импорт модуля файл не создаёт."""
    path = path or FILE_PATH
    wb = Workbook()
    wb.properties.created = FIXED_TIME
    wb.properties.modified = FIXED_TIME

    # ---------- Лист 1: все задачи ----------
    ws1 = wb.active
    ws1.title = "Все задачи"
    ws1.append(["ID", "Задача", "Приоритет", "Срок", "Статус"])
    for id_, title, pr, due, done in TASKS:
        ws1.append([id_, title, PRIORITIES[pr], fmt(due), STATUS[done]])

    for row in ws1.iter_rows(min_row=2, max_row=ws1.max_row, max_col=5):
        for c in row:
            c.border = border
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if c.column == 3:
                c.fill = priority_fill(c.value)

    for row in ws1.iter_rows(min_row=1, max_row=1, max_col=5):
        for c in row:
            c.fill = header_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center")

    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = f"A1:E{ws1.max_row}"
    for col, width in {"A": 6, "B": 42, "C": 12, "D": 12, "E": 12}.items():
        ws1.column_dimensions[col].width = width

    # ---------- Лист 2: план на неделю ----------
    ws2 = wb.create_sheet("План на неделю")
    ws2.append(["ID", "Задача", "Приоритет", "Срок", "Просрочка, дней", "Статус"])

    overdue = sorted(
        (t for t in TASKS if not t[4] and t[3] is not None and date.fromisoformat(t[3]) < TODAY),
        key=lambda t: (-(TODAY - date.fromisoformat(t[3])).days, PRIORITY_ORDER[t[2]]),
    )
    this_week = sorted(
        (t for t in TASKS
         if not t[4] and t[3] is not None and TODAY <= date.fromisoformat(t[3]) <= WEEK_END),
        key=lambda t: (date.fromisoformat(t[3]), PRIORITY_ORDER[t[2]]),
    )

    def write_block(rows, title, gap_label, gap_value):
        ws2.append([title])
        for row in ws2.iter_rows(min_row=ws2.max_row, max_col=6):
            for c in row:
                c.fill = alt_fill
                c.font = Font(bold=True)
                c.border = border
        ws2.append(["ID", "Задача", "Приоритет", "Срок", gap_label, "Статус"])
        for id_, title_, pr, due, done in rows:
            ws2.append([id_, title_, PRIORITIES[pr], fmt(due), gap_value(due), STATUS[done]])

    write_block(overdue, "ПРОСРОЧЕННЫЕ ЗАДАЧИ (до %s)" % fmt(TODAY), "Просрочка, дней",
                lambda due: (TODAY - date.fromisoformat(due)).days)
    write_block(this_week, "ЗАДАЧИ НА ТЕКУЩЕЙ НЕДЕЛЕ (%s – %s)" % (fmt(WEEK_START), fmt(WEEK_END)),
                "Дней до срока", lambda due: (date.fromisoformat(due) - TODAY).days)

    for row in ws2.iter_rows(min_row=1, max_row=ws2.max_row, max_col=6):
        for c in row:
            c.border = border
            if c.row > 1 and c.column == 3 and c.value in PRIORITIES.values():
                c.fill = priority_fill(c.value)
            if c.row > 1:
                c.alignment = Alignment(vertical="top", wrap_text=True)

    for col, width in {"A": 6, "B": 42, "C": 12, "D": 12, "E": 16, "F": 12}.items():
        ws2.column_dimensions[col].width = width

    save_workbook(wb, path)
    print(f"Сохранено: {path}")
    print(f"Всего задач: {len(TASKS)}, выполнено: {sum(t[4] for t in TASKS)}")
    print(f"Просрочено: {len(overdue)}, на текущей неделе: {len(this_week)}")


if __name__ == "__main__":
    build_report()