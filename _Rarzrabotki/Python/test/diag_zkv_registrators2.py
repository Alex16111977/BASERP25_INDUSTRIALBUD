# -*- coding: utf-8 -*-
"""Диагностика 2: реальные типы регистраторов ЗарплатаКВыплате и А_ВзСС (Постернак).
Классификация типов — на стороне Python через Метаданные().Имя."""
import sys
from collections import defaultdict
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

DRFO = "2742610332"  # Постернак Андрій Володимирович

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

def type_name(ref):
    try:
        return ref.Метаданные().Имя
    except Exception:
        try:
            return erp.XMLТипЗнч(ref).ИмяТипа
        except Exception:
            return "<?>"

def breakdown(register, resource, period_filter, fl_field, drfo=None, label=""):
    where = period_filter
    if drfo:
        where += f' И Д.{fl_field}.КодПоДРФО = "{drfo}"'
    q = erp.NewObject("Запрос")
    q.Text = f"""ВЫБРАТЬ
	Д.Регистратор КАК Рег,
	Д.ВидДвижения КАК ВД,
	СУММА(Д.{resource}) КАК Сумма
ИЗ
	РегистрНакопления.{register} КАК Д
ГДЕ
	{where}
СГРУППИРОВАТЬ ПО
	Д.Регистратор,
	Д.ВидДвижения"""
    t = q.Execute().Выгрузить()
    agg = defaultdict(lambda: [0.0, 0])
    for i in range(t.Количество()):
        r = t.Получить(i)
        vd = erp.XMLСтрока(r.ВД)
        key = (type_name(r.Рег), vd)
        agg[key][0] += float(r.Сумма or 0)
        agg[key][1] += 1
    print(f"--- {label} ({register}) ---")
    for (tn, vd), (s, n) in sorted(agg.items()):
        print(f"  {tn:45s} {vd:8s} Σ={s:,.2f} (движ.{n})")

# 1) ЗКВ: типы регистраторов, вся база за май 2026 (без ФЛ-фильтра)
breakdown("ЗарплатаКВыплате", "СуммаКВыплате",
          "Д.Период МЕЖДУ ДАТАВРЕМЯ(2026, 5, 1) И ДАТАВРЕМЯ(2026, 5, 31, 23, 59, 59)",
          "ФизическоеЛицо", None, "ЗКВ май 2026, вся база")

# 2) ЗКВ: Постернак декабрь 2025
breakdown("ЗарплатаКВыплате", "СуммаКВыплате",
          "Д.Период МЕЖДУ ДАТАВРЕМЯ(2025, 12, 1) И ДАТАВРЕМЯ(2025, 12, 31, 23, 59, 59)",
          "ФизическоеЛицо", DRFO, "ЗКВ дек 2025, Постернак")

# 3) А_ВзСС: Постернак по месяцам — кто пишет, ищем ОЗФУ=200000
for (y, m, d) in ((2025, 12, 31), (2026, 1, 31), (2026, 2, 28), (2026, 3, 31), (2026, 4, 30), (2026, 5, 31)):
    breakdown("А_ВзаиморасчетыССотрудниками", "СуммаВзаиморасчетов",
              f"Д.Период МЕЖДУ ДАТАВРЕМЯ({y}, {m}, 1) И ДАТАВРЕМЯ({y}, {m}, {d}, 23, 59, 59)",
              "ФизическоеЛицо", DRFO, f"А_ВзСС {y}-{m:02d}, Постернак")

print("DIAG2 DONE")
