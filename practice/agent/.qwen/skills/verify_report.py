#!/usr/bin/env python3
"""Проверка отчёта make_report.py: листы и данные — против исходного списка TASKS.

Запуск: python3 verify_report.py
Выводит OK при полном совпадении; при расхождении завершается с ошибкой.
"""
import hashlib
import os
import subprocess
import sys
from datetime import date, timedelta

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import make_report as mr

TODAY = date(2026, 10, 31)
WEEK_START = TODAY - timedelta(days=TODAY.weekday())  # понедельник текущей недели
WEEK_END = WEEK_START + timedelta(days=6)             # воскресенье текущей недели


def run_script():
    subprocess.run([sys.executable, os.path.join(HERE, "make_report.py")],
                   check=True, capture_output=True, text=True)


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        h.update(f.read())
    return h.hexdigest()


def overdue_rows_of(tasks):
    rows = [t for t in tasks if not t[4] and t[3] is not None and date.fromisoformat(t[3]) < TODAY]
    rows.sort(key=lambda t: (-(TODAY - date.fromisoformat(t[3])).days, mr.PRIORITY_ORDER[t[2]]))
    return [(t[0], t[1], mr.PRIORITIES[t[2]], mr.fmt(t[3]),
             (TODAY - date.fromisoformat(t[3])).days, "В работе") for t in rows]


def week_rows_of(tasks):
    rows = [t for t in tasks if not t[4] and t[3] is not None
            and TODAY <= date.fromisoformat(t[3]) <= WEEK_END]
    rows.sort(key=lambda t: (date.fromisoformat(t[3]), mr.PRIORITY_ORDER[t[2]]))
    return [(t[0], t[1], mr.PRIORITIES[t[2]], mr.fmt(t[3]),
             (date.fromisoformat(t[3]) - TODAY).days, "В работе") for t in rows]


def check_sheet1(ws1):
    assert [ws1.cell(row=1, column=c).value for c in range(1, 6)] == \
        ["ID", "Задача", "Приоритет", "Срок", "Статус"]
    assert ws1.max_row == len(mr.TASKS) + 1 == 64, (ws1.max_row, len(mr.TASKS))
    for idx, (i, title, pr, due, done) in enumerate(mr.TASKS, start=2):
        got = [ws1.cell(row=idx, column=c).value for c in range(1, 6)]
        want = [i, title, mr.PRIORITIES[pr], mr.fmt(due) if due else "—", mr.STATUS[done]]
        assert got == want, (idx, got, want)
    assert ws1.freeze_panes == "A2", ws1.freeze_panes
    assert ws1.auto_filter.ref == f"A1:E{len(mr.TASKS) + 1}", ws1.auto_filter.ref
    header = ws1.cell(row=1, column=1)
    assert header.font.bold is True and header.fill.patternType == "solid"
    assert ws1.cell(row=2, column=2).alignment.wrap_text is True


def check_sheet2(ws2):
    raw = list(ws2.iter_rows(values_only=True))
    find_blocks = lambda prefix: [i for i, r in enumerate(raw)
                                  if r and isinstance(r[0], str) and r[0].startswith(prefix)]
    b1, b2 = find_blocks("ПРОСРОЧЕННЫЕ")[0], find_blocks("ЗАДАЧИ НА ТЕКУЩЕЙ")[0]

    assert raw[b1][0] == "ПРОСРОЧЕННЫЕ ЗАДАЧИ (до %s)" % mr.fmt(TODAY), raw[b1][0]
    assert raw[b2][0] == "ЗАДАЧИ НА ТЕКУЩЕЙ НЕДЕЛЕ (%s – %s)" % (mr.fmt(WEEK_START), mr.fmt(WEEK_END)), raw[b2][0]
    assert [raw[b1 + 1][c] for c in range(6)] == \
        ["ID", "Задача", "Приоритет", "Срок", "Просрочка, дней", "Статус"]
    assert [raw[b2 + 1][c] for c in range(6)] == \
        ["ID", "Задача", "Приоритет", "Срок", "Дней до срока", "Статус"]

    overdue_rows = [tuple(r) for r in raw[b1 + 2:b2]]
    week_rows = [tuple(r) for r in raw[b2 + 2:]]
    assert overdue_rows == overdue_rows_of(mr.TASKS), list(zip(overdue_rows, overdue_rows_of(mr.TASKS)))
    assert week_rows == week_rows_of(mr.TASKS), list(zip(week_rows, week_rows_of(mr.TASKS)))
    assert len({r[0] for r in overdue_rows}) == len(overdue_rows)
    assert len({r[0] for r in week_rows}) == len(week_rows)

    assert ws2.cell(row=b1 + 1, column=1).font.bold is True
    assert ws2.cell(row=b1 + 1, column=1).fill.patternType == "solid"
    assert ws2.cell(row=b1 + 3, column=2).alignment.wrap_text is True


def main():
    run_script()
    assert os.path.isfile(mr.FILE_PATH), mr.FILE_PATH
    first = sha256(mr.FILE_PATH)
    run_script()
    assert sha256(mr.FILE_PATH) == first, "Перезапуск make_report.py даёт другой байтовый результат"

    wb = load_workbook(mr.FILE_PATH)
    assert wb.sheetnames == ["Все задачи", "План на неделю"], wb.sheetnames
    check_sheet1(wb["Все задачи"])
    check_sheet2(wb["План на неделю"])
    print("OK")


if __name__ == "__main__":
    main()