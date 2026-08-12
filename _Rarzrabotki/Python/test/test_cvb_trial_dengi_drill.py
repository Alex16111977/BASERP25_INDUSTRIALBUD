# -*- coding: utf-8 -*-
"""Дорасследование пробного прогона: из чего состоит остаток расхождений «Деньги безнал».
Read-only: фазы 1-2, разбор строк — счета (имена), непроведённые документы (суммы),
статусы проблемных строк без действия."""
import datetime
import sys
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

PLUGIN = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьДеньги.epf"
D1 = datetime.datetime(2026, 7, 1)
D2 = datetime.datetime(2026, 7, 19)

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

p = erp.ВнешниеОбработки.Создать(PLUGIN, False)
p.НачалоПериода = D1
p.ОкончаниеПериода = D2
p.СравнитьОстатки()

print("СЧЕТА С РАСХОЖДЕНИЕМ (имя | НачЕРП | НачБух | КонЕРП | КонБух | Δ):")
for i in range(p.ТаблицаРасхождений.Количество()):
    r = p.ТаблицаРасхождений.Получить(i)
    d = float(r.Разница)
    if abs(d) < 0.005:
        continue
    print("  %-52s | %14.2f | %14.2f | %14.2f | %14.2f | %14.2f"
          % (S(r.БанковскийСчет)[:52], float(r.НачОстатокЕРП), float(r.НачОстатокБух),
             float(r.ОстатокЕРП), float(r.ОстатокБух), d))

p.АнализироватьДокументы(-1)

print()
print("НЕПРОВЕДЁННЫЕ В ЕРП (действие «Провести в ЕРП»):")
tot = 0.0
for i in range(p.ТаблицаДокументов.Количество()):
    r = p.ТаблицаДокументов.Получить(i)
    if str(r.Действие).strip() != "Провести в ЕРП":
        continue
    dl = float(r.СуммаБух) - float(r.СуммаЕРП)
    tot += dl
    print("  %-46s | счёт %-32s | ЕРП %12.2f | Бух %12.2f"
          % (S(r.ДокументЕРП)[:46], S(r.БанковскийСчет)[:32], float(r.СуммаЕРП), float(r.СуммаБух)))
print("  Σ (Бух-ЕРП) по непроведённым: %.2f" % tot)

print()
print("ПРОБЛЕМНЫЕ СТРОКИ БЕЗ ДЕЙСТВИЯ (ЕстьРасхождение=Истина, действие пусто) — по статусам:")
by_status = {}
for i in range(p.ТаблицаДокументов.Количество()):
    r = p.ТаблицаДокументов.Получить(i)
    if not bool(r.ЕстьРасхождение) or str(r.Действие).strip():
        continue
    st = str(r.Статус).strip()
    key = st[:60]
    ent = by_status.setdefault(key, [0, 0.0])
    ent[0] += 1
    ent[1] += float(r.СуммаЕРП) - float(r.СуммаБух)
for st, (cnt, dsum) in sorted(by_status.items(), key=lambda x: -x[1][0]):
    print("  %-60s : %4d строк | Σ(ЕРП-Бух) %14.2f" % (st, cnt, dsum))
print("DRILL DONE")
