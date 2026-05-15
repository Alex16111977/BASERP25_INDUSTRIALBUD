# -*- coding: utf-8 -*-
"""ДІАГНОСТИКА: класифікація 16 статей за реквізитом АктивПассив
(як штатний Отчёт.УправленческийБаланс). 2 кроки + join у Python. Read-only."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom
from datetime import datetime

CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
D1 = datetime(2026, 2, 1, 12, 0, 0)
D2 = datetime(2026, 2, 28, 23, 59, 59)

pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(CONN)

q0 = conn.NewObject("Запрос")
q0.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ Справочник.Организации ГДЕ КодПоЕДРПОУ = "40645273"'
s = q0.Выполнить().Выбрать(); s.Следующий(); org = s.С

def k(ref):
    try: return str(conn.XMLСтрока(ref))
    except Exception: return repr(ref)

# Крок А: довідник статей з типом АктивПассив (пряма таблиця ПВХ)
qa = conn.NewObject("Запрос")
qa.Текст = """ВЫБРАТЬ Ссылка КАК С, Наименование КАК Н,
ПРЕДСТАВЛЕНИЕ(АктивПассив) КАК Тип
ИЗ ПланВидовХарактеристик.СтатьиАктивовПассивов ГДЕ НЕ ЭтоГруппа"""
tza = qa.Выполнить().Выгрузить()
типы = {}
for i in range(tza.Количество()):
    r = tza.Получить(i)
    типы[k(r.С)] = (str(r.Н), str(r.Тип))
print(f"[A] статей у ПВХ (не групи): {tza.Количество()}")

# Крок Б: ПАП.ОстаткиИОбороты по статті (без join, без виключення)
qb = conn.NewObject("Запрос")
qb.Текст = """ВЫБРАТЬ Б.Статья КАК Статья, ПРЕДСТАВЛЕНИЕ(Б.Статья) КАК Н,
СУММА(Б.СуммаКонечныйОстаток) КАК sK
ИЗ РегистрНакопления.ПрочиеАктивыПассивы.ОстаткиИОбороты(&Д1,&Д2,Авто,,Организация=&Орг) КАК Б
СГРУППИРОВАТЬ ПО Б.Статья, ПРЕДСТАВЛЕНИЕ(Б.Статья)
ИМЕЮЩИЕ СУММА(Б.СуммаКонечныйОстаток) <> 0"""
qb.УстановитьПараметр("Д1", D1); qb.УстановитьПараметр("Д2", D2); qb.УстановитьПараметр("Орг", org)
tzb = qb.Выполнить().Выгрузить()

by_type = {}
print(f"\n{'Стаття':<40}{'Тип':<16}{'КонОст':>18}")
print("-" * 74)
rows = []
for i in range(tzb.Количество()):
    r = tzb.Получить(i)
    key = k(r.Статья); nm = str(r.Н); v = float(r.sK or 0)
    tp = типы.get(key, ("?", "?НЕ_В_ПВХ"))[1]
    rows.append((nm, tp, v))
for nm, tp, v in sorted(rows, key=lambda x: (x[1], -x[2])):
    by_type.setdefault(tp, 0.0); by_type[tp] += v
    print(f"{nm[:38]:<40}{tp:<16}{v:>18,.2f}")

print("-" * 74)
print("Σ за типом АктивПассив:")
for tp, v in sorted(by_type.items()):
    print(f"  {tp:<22} = {v:>18,.2f}")
akt = by_type.get("Актив", 0.0)
pas = by_type.get("Пассив", 0.0)
ap  = by_type.get("АктивПассив", 0.0)
print(f"\nАктив={akt:,.2f}  Пассив={pas:,.2f}  АктивПассив={ap:,.2f}")
print(f"Σ повний набір = {akt+pas+ap:,.2f}  (≈0 → справжній баланс)")
print(f"Контроль штатного звіту: Актив-секція = |Пассив+АктивПассив|?")
print(f"  Актив = {akt:,.2f}")
print(f"  |Пассив+АктивПассив| = {abs(pas+ap):,.2f}")
print(f"  Δ = {akt - abs(pas+ap):,.2f}")
