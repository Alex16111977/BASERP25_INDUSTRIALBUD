# -*- coding: utf-8 -*-
"""Репліка клієнтського парсера форми (ПрочитатьКнигуExcel) — 1:1 логіка BSL.

Читає книгу через той самий Excel COM, тими самими правилами:
  рядок 3 — заголовки місяців у колонках I(9)/K(11)/M(13);
  дані з 5-го рядка; B(2) — назва позиції; кіл = колонка місяця, сума = наступна.

Мета — переконатися, що парсер дає рівно те, що очікуємо, ДО натискання кнопки в UI.
"""
import win32com.client, sys, json
sys.stdout.reconfigure(encoding='utf-8')

PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\А_ПланФактныйПроизводствоПолный\Бюджет\МАТЕРИАЛИ.xlsx"

MONTHS = [
    (("січ", "янв"), 1), (("лют", "фев"), 2), (("берез", "март"), 3),
    (("квіт", "квит", "апрел"), 4), (("трав", "май"), 5), (("черв", "июн"), 6),
    (("лип", "июл"), 7), (("серп", "август"), 8), (("верес", "сентябр"), 9),
    (("жовт", "октябр"), 10), (("листоп", "ноябр"), 11), (("груд", "декабр"), 12),
]


def month_from_header(text, year=2026):
    t = (text or "").strip().lower()
    if not t:
        return None
    for keys, num in MONTHS:
        if any(k in t for k in keys):
            return "%04d-%02d-01" % (year, num)
    return None


def num(v):
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except Exception:
        return 0.0


excel = win32com.client.Dispatch("Excel.Application")
excel.DisplayAlerts = False
excel.Visible = False
rows, sheets_info = [], []
err = ""
book = None
try:
    book = excel.Workbooks.Open(PATH, 0, True)
    for i in range(1, book.Worksheets.Count + 1):
        sh = book.Worksheets(i)
        name = str(sh.Name).strip()
        cols = {}
        for c in (9, 11, 13):
            m = month_from_header(sh.Cells(3, c).Text)
            if m:
                cols[c] = m
        try:
            last = sh.Cells(1, 1).SpecialCells(11).Row
        except Exception:
            last = 0
        cnt = 0
        for r in range(5, int(last) + 1):
            nm = str(sh.Cells(r, 2).Text or "").strip()
            if not nm:
                continue
            for c, mon in cols.items():
                kol = num(sh.Cells(r, c).Value)
                summ = num(sh.Cells(r, c + 1).Value)
                if kol == 0 and summ == 0:
                    continue
                rows.append({"Лист": name, "НоменклатураСС": nm, "Месяц": mon,
                             "Количество": kol, "Сумма": round(summ, 2)})
                cnt += 1
        sheets_info.append((name, int(last), sorted(cols.values()), cnt,
                            round(sum(x["Сумма"] for x in rows if x["Лист"] == name), 2)))
except Exception as e:
    err = str(e)
finally:
    try:
        if book is not None:
            book.Close(False)
    except Exception:
        pass
    try:
        excel.Quit()
    except Exception:
        pass
    excel = None

if err:
    print("ПОМИЛКА:", err)
    sys.exit(1)

print("=== ЩО ПРОЧИТАЄ ФОРМА З УСІЄЇ КНИГИ ===")
print("%-16s%9s  %-26s%9s%16s" % ("Аркуш", "рядків", "місяці", "записів", "сума"))
for n, last, mons, cnt, s in sheets_info:
    print("%-16s%9d  %-26s%9d%16s" % (n[:14], last, ",".join(m[:7] for m in mons), cnt,
                                      "{:,.2f}".format(s)))
print("-" * 80)
print("%-16s%9s  %-26s%9d%16s" % ("РАЗОМ", "", "", len(rows),
                                  "{:,.2f}".format(sum(x["Сумма"] for x in rows))))

irs = [x for x in rows if x["Лист"] in ("IRS 15", "IRS 30")]
other = [x for x in rows if x["Лист"] not in ("IRS 15", "IRS 30")]
print("\n  з них IRS 15/30 : %5d записів на %s" % (len(irs), "{:,.2f}".format(sum(x["Сумма"] for x in irs))))
print("  ЧУЖІ аркуші     : %5d записів на %s" % (len(other), "{:,.2f}".format(sum(x["Сумма"] for x in other))))
print("\n  еталон графіка IRS: 338 записів на 11 286 390,58")

json.dump(rows, open("data_form_replica.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("\n-> data_form_replica.json")
