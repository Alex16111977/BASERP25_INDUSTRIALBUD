# -*- coding: utf-8 -*-
"""
СКРИПТ 05 — Drill наличных per (Подр, Касса) — зеркало штатной Ведомости

ARGV: <YYYY-MM>
Артефакт: _artifacts/05_nalich_<месяц>.csv
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp, get_refs, money, save_csv

if len(sys.argv) < 2:
    print("Usage: python 05_drill_nalich_per_kassa.py <YYYY-MM>"); sys.exit(1)
y, m = map(int, sys.argv[1].split("-"))
end_y = y if m < 12 else y + 1
end_m = m + 1 if m < 12 else 1

erp = connect_erp()
refs = get_refs(erp)
S = erp.String

print("=" * 110)
print(f"СКРИПТ 05 — Наличные per (Подр, Касса) за {sys.argv[1]}")
print("=" * 110)

q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", refs["Орг"])
q.Текст = f"""
ВЫБРАТЬ
    ЕСТЬNULL(Н.Подразделение, ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)) КАК Подр,
    Н.Касса КАК Касса,
    СУММА(Н.СуммаНачальныйОстаток) КАК НО,
    СУММА(Н.СуммаПриход) КАК Приход,
    СУММА(Н.СуммаРасход) КАК Расход,
    СУММА(Н.СуммаКонечныйОстаток) КАК КМ
ИЗ РегистрНакопления.ДенежныеСредстваНаличные.ОстаткиИОбороты(
    ДАТАВРЕМЯ({y},{m},1,0,0,0), КОНЕЦПЕРИОДА(ДАТАВРЕМЯ({y},{m},15),МЕСЯЦ), , , Организация = &Орг) КАК Н
СГРУППИРОВАТЬ ПО ЕСТЬNULL(Н.Подразделение, ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)), Н.Касса
ИМЕЮЩИЕ СУММА(Н.СуммаКонечныйОстаток) <> 0 ИЛИ СУММА(Н.СуммаНачальныйОстаток) <> 0
УПОРЯДОЧИТЬ ПО Подр, КМ УБЫВ
"""
r = q.Выполнить().Выгрузить()
print(f"\nСтрок: {r.Количество()}")

print(f"\n{'Подр':<25}{'Касса':<40}{'НО':>15}{'Приход':>15}{'Расход':>15}{'КМ':>15}")
print("-" * 125)
rows = []
for i in range(r.Количество()):
    rec = r.Получить(i)
    podr = str(S(rec.Подр) or "(пусто)")
    k = str(S(rec.Касса) or "")
    no = float(rec.НО or 0); pr = float(rec.Приход or 0); rs = float(rec.Расход or 0); km = float(rec.КМ or 0)
    print(f"{podr[:23]:<25}{k[:38]:<40}{money(no):>15}{money(pr):>15}{money(rs):>15}{money(km):>15}")
    rows.append({"Подр": podr, "Касса": k, "НО": no, "Приход": pr, "Расход": rs, "КМ": km})

save_csv(f"05_nalich_{sys.argv[1]}", rows, ["Подр", "Касса", "НО", "Приход", "Расход", "КМ"])

# Pivot
print("\n--- Касса в нескольких Подр ---")
by_k = {}
for r in rows:
    by_k.setdefault(r["Касса"], []).append(r)
for k, lst in by_k.items():
    if len(lst) > 1:
        print(f"  Касса: {k}")
        for x in lst:
            print(f"     Подр={x['Подр'][:30]:<32}  КМ={money(x['КМ']):>15}")
print(f"\nАртефакт: 05_nalich_{sys.argv[1]}.csv")
