# -*- coding: utf-8 -*-
"""
СКРИПТ 18 — Перепроведение документов А_ФинРез_Баланс

Параметр (env / argv): месяц или ALL
  python 18_repost_finrez_balans.py 2025-12   — только декабрь 2025
  python 18_repost_finrez_balans.py ALL       — все 5 за период

Перепроведение пересобирает А_ОтчетБаланс_Свод (на этом документ-агрегаторе
держится сводный регистр).
"""
import sys, io, os, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from datetime import datetime, timedelta
from _common import connect_erp, get_refs

erp = connect_erp()
refs = get_refs(erp)
ORG = refs["Орг"]
РежимПроведения = erp.РежимЗаписиДокумента.Проведение

# argv: месяц вида YYYY-MM или ALL
target = sys.argv[1] if len(sys.argv) > 1 else "ALL"
print(f"СКРИПТ 18 — перепроведение А_ФинРез_Баланс (target={target})")
print(f"Старт: {datetime.now().strftime('%H:%M:%S')}\n")

q = erp.NewObject("Запрос")
q.УстановитьПараметр("Орг", ORG)
q.УстановитьПараметр("Н1", datetime(2025, 12, 1, 0, 0, 0))
q.УстановитьПараметр("Н2", datetime(2026, 12, 31, 23, 59, 59))
q.Текст = """
ВЫБРАТЬ
    Ссылка, Номер, Дата, Месяц, Проведен
ИЗ Документ.А_ФинРез_Баланс
ГДЕ Организация = &Орг
    И Месяц МЕЖДУ &Н1 И &Н2
    И ВключатьДочерние = ЛОЖЬ
    И НЕ ПометкаУдаления
УПОРЯДОЧИТЬ ПО Месяц
"""
r = q.Выполнить().Выгрузить()

target_docs = []
for i in range(r.Количество()):
    rec = r.Получить(i)
    мес_str = rec.Месяц.strftime("%Y-%m") if hasattr(rec.Месяц, "strftime") else str(rec.Месяц)[:7]
    if target == "ALL" or мес_str == target:
        target_docs.append((rec.Ссылка, rec.Номер, rec.Дата, мес_str))

print(f"К перепроведению: {len(target_docs)} документов")
for ref, ном, дата, мес in target_docs:
    print(f"  №{ном} от {дата} (Месяц={мес})")

ok = err = 0
t0 = time.time()
for ref, ном, дата, мес in target_docs:
    t1 = time.time()
    try:
        obj = ref.ПолучитьОбъект()
        if obj is None:
            print(f"  №{ном} {мес}: ERR PolychitObj=None"); err += 1; continue
        obj.Записать(РежимПроведения)
        dt1 = time.time() - t1
        print(f"  №{ном} {мес}: OK за {dt1:.1f}s")
        ok += 1
    except Exception as e:
        msg = str(e)[:200]
        if hasattr(e, "excepinfo") and e.excepinfo and e.excepinfo[2]:
            msg = str(e.excepinfo[2])[:200]
        print(f"  №{ном} {мес}: ERR {msg}")
        err += 1

dt_total = time.time() - t0
print(f"\nИТОГО за {int(dt_total)}s: OK={ok}, ERR={err}")
sys.exit(0 if err == 0 else 1)
