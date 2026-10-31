#!/usr/bin/env python3
"""Верификация отчёта: листы и данные против исходного списка задач."""
import sys
from datetime import date
from openpyxl import load_workbook

sys.path.insert(0, "/home/monkler/aidd_week_04_mcp/practice/agent")
import make_report as mr

TODAY = date(2026, 10, 31)
WEEK_END = date(2026, 11, 1)

wb = load_workbook(mr.FILE_PATH)
assert wb.sheetnames == ["Все задачи", "План на неделю"], wb.sheetnames

# --- Лист 1 ---
ws1 = wb["Все задачи"]
got1 = set()
for r in ws1.iter_rows(min_row=2, max_col=5, values_only=True):
    got1.add((r[0], r[1], r[2], r[3], r[4]))
expected1 = set(
    (i, title, mr.PRIORITIES[pr], mr.fmt(due) if due else "—", mr.STATUS[done])
    for i, title, pr, due, done in mr.TASKS
)
assert got1 == expected1, f"Лист 1 mismatch: {got1 ^ expected1} "
assert ws1.max_row - 1 == len(mr.TASKS) == 63, ws1.max_row
print("Лист 1 OK: 63 задачи, наборы совпадают")

# --- Лист 2 ---
ws2 = wb["План на неделю"]
raw = list(ws2.iter_rows(min_row=2, values_only=True))

def block_starts(rows, prefix):
    return [i for i, r in enumerate(rows) if isinstance(r[0], str) and r[0].startswith(prefix)]

overdue_rows = raw[block_starts(raw, "ПРОСРОЧЕННЫЕ")[0] + 2:block_starts(raw, "ЗАДАЧИ НА ТЕКУЩЕЙ")[0]]
week_rows = raw[block_starts(raw, "ЗАДАЧИ НА ТЕКУЩЕЙ")[0] + 2:]

def parse_block(rows):
    out = {}
    for r in rows:
        id_, title, pr, due_s, gap, status = r
        out[id_] = (title, pr, due_s, gap, status)
    return out

ov = parse_block(overdue_rows)
wk = parse_block(week_rows)

exp_overdue = {
    t[0]: (t[1], mr.PRIORITIES[t[2]], mr.fmt(t[3]), (TODAY - date.fromisoformat(t[3])).days, "В работе")
    for t in mr.TASKS if not t[4] and t[3] is not None and date.fromisoformat(t[3]) < TODAY
}
exp_week = {
    t[0]: (t[1], mr.PRIORITIES[t[2]], mr.fmt(t[3]), (date.fromisoformat(t[3]) - TODAY).days, "В работе")
    for t in mr.TASKS if not t[4] and t[3] is not None and TODAY <= date.fromisoformat(t[3]) <= WEEK_END
}

assert set(ov) == set(exp_overdue), f"Просроченные: {set(ov) ^ set(exp_overdue)}"
for k in exp_overdue:
    assert ov[k] == exp_overdue[k], (k, ov[k], exp_overdue[k])
# порядок: по просрочке убыв., при равенстве приоритет по возрастанию
exp_overdue_sorted = [
    t[0] for t in sorted(
        (t for t in mr.TASKS if not t[4] and t[3] is not None and date.fromisoformat(t[3]) < TODAY),
        key=lambda t: ((TODAY - date.fromisoformat(t[3])).days, -mr.PRIORITY_ORDER[t[2]]),
    )
]
assert [r[0] for r in overdue_rows] == exp_overdue_sorted, (
    [r[0] for r in overdue_rows], exp_overdue_sorted)
print(f"Лист 2 OK: просроченных {len(ov)}, на текущей неделе {len(wk)} (порядок по просрочке)")
assert set(wk) == set(exp_week), f"Неделя: {set(wk) ^ set(exp_week)}"
for k in exp_week:
    assert wk[k] == exp_week[k], (k, wk[k], exp_week[k])

print("Проверки пройдены: файл", mr.FILE_PATH)