# -*- coding: utf-8 -*-
"""План Факт Производство 09-10-2026 + сценарии (как образец 18.09 «новый_исправленный»).
Колонки A–O — выгрузка казны без изменений; P–W — сценарии с живыми формулами; низ — Ф1/Ф2/ФА/Free CF/Fix."""
import sys, os, json, copy
sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from dump import fixed_copy
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter

BASE = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Письмо Директору по отчетам на производство\09102026"
SRC = os.path.join(BASE, "План Факт Производство 09-10-2026.xlsx")
OUT = os.path.join(BASE, "План Факт Производство 09-10-2026 сценарии.xlsx")
from fo_docs import read as read_fo
FO = read_fo()


def fo(prefix):
    """opt — «ИТОГО к получению» (K); pess — подписанные акты со сроком оплаты до конца месяца."""
    for k, v in FO.items():
        if k.startswith(prefix):
            return {"opt": v["K"], "pess": v["pess"]}
    raise KeyError(prefix)


wb = openpyxl.load_workbook(fixed_copy(SRC))
ws = wb["TDSheet"]
rows = {}
for r in range(6, ws.max_row + 1):
    a = ws["A%d" % r].value
    if a:
        rows[str(a).strip()] = r
TOT = rows["Итого"]
FIRST, LAST = 6, TOT - 1


def num(c):
    v = ws[c].value
    return float(v) if isinstance(v, (int, float)) else 0.0


# ---- вводные по объектам: (опт, песс, потребность|None, особая строка, пояснение) ----
L = lambda name: num("L%d" % rows[name])  # остаток плана затрат месяца
D = lambda name: num("D%d" % rows[name])
E = lambda name: num("E%d" % rows[name])
inp = {}
k = fo("МК Кривий ріг")
inp["МК Кривий ріг"] = (k["opt"], k["pess"], None, False,
    "финотчёт 09.10: к получению 422 766,20 (в т.ч. не подписанные / не проведённые акты 294 158,90); песс — 3 подписанных акта от 18.09, срок оплаты 25.09 прошёл, просрочено 14 дн. — 175 243,05")
k = fo("МК Каменец Цемент")
inp["МК Каменец Цемент"] = (k["opt"], 0.0, None, False,
    "финотчёт: к получению 629 502,80 по актам на рассмотрении 1 259 005,60 (за вычетом аванса); песс 0 — акты не подписаны")
k = fo("МД МХП ОРІЛЬ")
inp["МД МХП ОРІЛЬ"] = (k["opt"], k["pess"], None, False,
    "финотчёт: к получению 419 894,73; песс — акт СТ00-000002 подписан 23.09, срок оплаты 03.10 прошёл (просрочено 6 дн.) — 316 611,41; акт 129 104,15 не подписан")
k = fo("МК Чорноморськ")
inp["МК Чорноморськ"] = (k["opt"], 0.0, None, False, "финотчёт: акт СТ00-000120 12 654 не проведён → опт 10 545 (за вычетом НДС), песс 0")
k = fo("МК Глобино")
inp["МК Глобино"] = (k["opt"], 0.0, None, False, "финотчёт: к получению 32 005,12 только по поданным актам → песс 0")
k = fo("МД IRC 2026")
inp["МД IRC 2026"] = (k["opt"], 0.0, None, False,
    "финотчёт: 2 акта 30 м по 1 774 329,06 (ІБ00-000171 от 22.09, ІБ00-000178 от 01.10) не подписаны → опт 3 548 658,12, песс 0")
inp["МД ООН 2025"] = (D("МД ООН 2025"), 0.0, None, False,
    "гарантийный платёж 1 154 400,18 (финотчёт: удержан, кол. I) — только в оптимистичном; песс 0: документов нет, месяц не платят (решение 09.10)")
inp["МД ООН Админ"] = (D("МД ООН Админ") - E("МД ООН Админ"), 0.0, None, False,
    "план поступлений октября 2 525 114 (по договору 20 % при поставке); актов нет → песс 0 (как 25.09); подтвердить")
inp["МД Кернел Кухня"] = (D("МД Кернел Кухня"), 0.0, None, False,
    "ДЦ на рассмотрении 716 773 (финотчёт, кол. D); оплата не подтверждена → песс 0")
inp["МД ООН 2026"] = (0.0, 0.0, None, True,
    "аванс получен ранее (6 498 439,86); затраты октября — из него: доп. маржа = −потребность (как 18.09); потребность производство не дало → 0")
inp["МХП санітарний вузол"] = (0.0, 0.0, None, False, "в финотчёте актов нет → поступлений не ставим (план 114 556)")
for name in rows:
    if name in ("Итого",):
        continue
    inp.setdefault(name, (0.0, 0.0, None, False, ""))
# потребность: заявки производства; на 09.10 производство их не дало — колонки R/V пустые (вписать, когда дадут)

# ---- шапка P–W ----
hdr4, hdr5 = ws["M4"], ws["M5"]
for c0, title in (("P", "Оптимистичный сценарий"), ("T", "Пессимистичный сценарий")):
    cell = ws["%s4" % c0]
    cell.value = title
    cell._style = copy.copy(hdr4._style)
ws.merge_cells("P4:S4")
ws.merge_cells("T4:W4")
for i, t in enumerate(["Поступления", "Плановый Лимит с учетом поступлений", "Потребность", "Дополнительная Маржа"] * 2):
    col = get_column_letter(16 + i)
    c = ws["%s5" % col]
    c.value = t
    c._style = copy.copy(hdr5._style)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for col in "PQRSTUVW":
    ws["%s4" % col]._style = copy.copy(hdr4._style)
ws.row_dimensions[5].height = 45
YELLOW = PatternFill("solid", fgColor="FFF2CC")


def lim(p, r):
    return "=IF({p}{r}+N{r}>0,IF({p}{r}<>0,IF(N{r}+{p}{r}>M{r},{p}{r}+O{r},0),0),0)".format(p=p, r=r)


calc = {}
for name, r in rows.items():
    if name == "Итого":
        continue
    p, t, need, prepaid, note = inp[name]
    for col in "PQRSTUVW":
        ws["%s%d" % (col, r)]._style = copy.copy(ws["O%d" % r]._style)
    ws["P%d" % r] = round(p, 2)
    ws["T%d" % r] = round(t, 2)
    ws["R%d" % r] = round(need, 2) if need else None
    ws["Q%d" % r] = lim("P", r)
    ws["U%d" % r] = lim("T", r)
    ws["V%d" % r] = None  # потребность пессимистичного — тоже вводная, производство не дало
    ws["S%d" % r] = ("=P{r}-Q{r}-R{r}" if prepaid else "=P{r}-Q{r}").format(r=r)
    ws["W%d" % r] = ("=T{r}-U{r}-V{r}" if prepaid else "=T{r}-U{r}").format(r=r)
    for col in "PQRSTUVW":
        ws["%s%d" % (col, r)].number_format = "#,##0.00"
    for col in "PTRV":
        ws["%s%d" % (col, r)].fill = YELLOW
    ws["X%d" % r] = note or None
    ws["X%d" % r].alignment = Alignment(wrap_text=True, vertical="top")
    # тот же расчёт в Python — для выводов и проверки
    M, N, O = num("M%d" % r), num("N%d" % r), num("O%d" % r)
    q = (p + O) if (p + N > 0 and p != 0 and N + p > M) else 0.0
    u = (t + O) if (t + N > 0 and t != 0 and N + t > M) else 0.0
    rr = need or 0.0
    calc[name] = dict(P=p, Q=q, R=rr, S=p - q - (rr if prepaid else 0), T=t, U=u, V=rr, W=t - u - (rr if prepaid else 0))
for col in "PQRSTUVW":
    ws["%s%d" % (col, TOT)] = "=SUM(%s%d:%s%d)" % (col, FIRST, col, LAST)
    ws["%s%d" % (col, TOT)]._style = copy.copy(ws["O%d" % TOT]._style)
    ws["%s%d" % (col, TOT)].number_format = "#,##0.00"
for col in "PQRSTUVW":
    ws.column_dimensions[col].width = 17
ws.column_dimensions["X"].width = 70
tot = {c: sum(v[c] for v in calc.values()) for c in "PQRSTUVW"}

# Нижний раздел (Ф1/Ф2/ФА, Free CF, Fix, выводы) по решению пользователя 09.10 не выводится.
from scen_notes import add_notes
add_notes(ws, rows, inp, calc, num, TOT, [])
wb.save(OUT)
print("saved", OUT)
for n, v in calc.items():
    if any(abs(x) > 0.004 for x in v.values()):
        print("%-28s P=%13.2f Q=%13.2f R=%13.2f S=%14.2f | T=%13.2f U=%13.2f W=%14.2f" % (n, v["P"], v["Q"], v["R"], v["S"], v["T"], v["U"], v["W"]))
print("ИТОГО  P=%.2f Q=%.2f R=%.2f S=%.2f | T=%.2f U=%.2f W=%.2f" % tuple(tot[c] for c in "PQRSTUW"))
