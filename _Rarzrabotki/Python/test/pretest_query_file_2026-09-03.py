# -*- coding: utf-8 -*-
"""Rule #-1: прогнать текст запроса из файла на BaseERP через COM и показать сводку.

Использование: python pretest_query_file_2026-09-03.py <файл .sql> [колонка для группового подсчёта]
Пакетные запросы (с ПОМЕСТИТЬ / ;) поддерживаются — результат последнего запроса пакета.
"""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

path = sys.argv[1]
col = sys.argv[2] if len(sys.argv) > 2 else None

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

q = erp.NewObject("Запрос")
q.Text = open(path, encoding="utf-8-sig").read()
try:
    r = q.Execute().Выгрузить()
except Exception as e:
    msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
    print("FAIL:", msg)
    sys.exit(1)

print(f"OK rows={r.Количество()} cols={r.Колонки.Количество()}")
names = [S(r.Колонки.Получить(i).Имя) for i in range(r.Колонки.Количество())]
print("cols:", ", ".join(names))
if r.Количество():
    x = r.Получить(0)
    print("row0:", " | ".join(f"{n}={S(getattr(x, n))}" for n in names))
if col:
    agg = {}
    for i in range(r.Количество()):
        v = S(getattr(r.Получить(i), col))
        agg[v] = agg.get(v, 0) + 1
    print(f"группировка по {col}:", dict(sorted(agg.items(), key=lambda kv: -kv[1])[:12]))
