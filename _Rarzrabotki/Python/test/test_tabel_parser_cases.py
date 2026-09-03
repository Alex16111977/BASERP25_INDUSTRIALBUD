# -*- coding: utf-8 -*-
"""Тест разбора ячейки табеля Казны: Документы.ТабельУчетаРабочегоВремени.РазобратьЗначениеЯчейки(v).

Запускать ПОСЛЕ загрузки модуля менеджера в kazna (/db-load-xml + /db-update).
Каждый случай печатается PASS/FAIL, exit code 1 при любом провале.
"""
import sys
import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

CONN = 'Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"'

# (ввод, Состояние, ЧасыФакт, ЧасыОфициальные, Ошибка пустая?)
CASES = [
    ("",      "",  0,    0, True),
    ("8",     "Р", 8,    8, True),     # число: ЧасыФакт и ЧасыОфициальные
    ("10,5",  "Р", 10.5, 10.5, True),
    ("2.4",   "Р", 2.4,  2.4, True),
    ("01",    "Р", 1,    1, True),
    ("8м",    "М", 0,    8, True),
    ("8М",    "М", 0,    8, True),
    ("8M",    "М", 0,    8, True),   # латинская M
    (" 8м ",  "М", 0,    8, True),
    ("М",     "М", 0,    0, True),
    ("м",     "М", 0,    0, True),
    ("О",     "О", 0,    0, True),
    ("Б",     "Б", 0,    0, True),
    ("Н",     "Н", 0,    0, True),
    ("В",     "В", 0,    0, True),
    ("С",     "С", 0,    0, True),
    ("К",     "К", 0,    0, True),
    ("Д",     "Д", 0,    0, True),
    ("O",     "О", 0,    0, True),   # латинская O
    ("1О",    "?", 0,    0, False),  # цифра + кириллическая О
    ("10-",   "?", 0,    0, False),
    ("8К",    "?", 0,    0, False),
    ("ОМ",    "?", 0,    0, False),
    ("Х",     "?", 0,    0, False),
]


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    kz = v8.Connect(CONN)
    mgr = kz.Документы.ТабельУчетаРабочегоВремени
    fails = 0
    for src, st, hf, ho, ok in CASES:
        r = mgr.РазобратьЗначениеЯчейки(src)
        got = (str(r.Состояние), float(r.ЧасыФакт), float(r.ЧасыОфициальные), str(r.Ошибка) == "")
        exp = (st, float(hf), float(ho), ok)
        status = "PASS" if got == exp else "FAIL"
        if status == "FAIL":
            fails += 1
        print(f"{status} {src!r:9} -> Состояние={got[0]!r} ЧасыФакт={got[1]} ЧасыОфиц={got[2]} ok={got[3]}"
              + ("" if status == "PASS" else f"   ОЖИДАЛОСЬ {exp}  err={r.Ошибка}"))
    print(f"\nИтого: {len(CASES)} случаев, провалов {fails}")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
