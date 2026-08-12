# -*- coding: utf-8 -*-
"""ЦВБ: финальный контрольный прогон «Тільки звіт» за 01.01–31.07.2026 (история целиком)."""
import datetime
import sys
import time
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String
engine = erp.А_ВыравниваниеБазСервер

nach = datetime.datetime(2026, 1, 1)
kon = datetime.datetime(2026, 7, 31)
session = datetime.datetime.now().replace(microsecond=0)

print(f"=== КОНТРОЛЬ ИСТОРИИ {nach:%d.%m.%Y} — {kon:%d.%m.%Y} (только отчёт) ===", flush=True)
for name in ("Деньги безнал", "Товары", "Взаиморасчёты", "Касса"):
    ref = erp.Справочники.А_КонтурыСверкиБаз.НайтиПоНаименованию(name, True)
    t0 = time.time()
    res = engine.ОбработатьКонтурПоЗначениям(ref, nach, kon, False, True, 500, 200, session, 1)
    m = res.Метрика
    print(f"[{name}] расходится={m.КлючейРасходится} объяснено={m.КлючейОбъяснено} "
          f"Σ|Δ|={float(m.СуммаДельта):,.2f} (объяснённых Σ={float(m.СуммаДельтаОбъяснено):,.2f}) "
          f"B={m.КПодтверждению}+{m.ПодтвержденийСверхЛимита} C={m.РучнойРазбор} "
          f"за {time.time()-t0:.0f} c", flush=True)

erp = None
print("FINAL CHECK DONE", flush=True)
