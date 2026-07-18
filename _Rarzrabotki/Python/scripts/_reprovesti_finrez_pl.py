# -*- coding: utf-8 -*-
"""Массовый перепровод всех А_ФинРез_PL за 2024-01..2026-04."""
import sys, io, json, os, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = conn.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ Ссылка, Номер, Дата, Месяц, Организация
ИЗ Документ.А_ФинРез_PL
ГДЕ Проведен И НЕ ПометкаУдаления
УПОРЯДОЧИТЬ ПО Месяц, Организация
"""
tz = q.Выполнить().Выгрузить()
print(f"Найдено А_ФинРез_PL к перепроводу: {tz.Количество()}\n")

mode = conn.PredefinedValue("РежимЗаписиДокумента.Проведение")
results = []
ok = fail = 0

for i in range(tz.Количество()):
    row = tz.Получить(i)
    obj = row.Ссылка.ПолучитьОбъект()
    if obj is None:
        results.append({"номер": str(row.Номер), "статус": "FAIL", "ошибка": "obj is None"})
        fail += 1
        print(f"  FAIL {row.Номер} {row.Месяц}: obj is None")
        continue
    try:
        obj.Записать(mode)
        results.append({
            "номер": str(row.Номер), "месяц": str(row.Месяц),
            "орг": conn.String(row.Организация), "статус": "OK"
        })
        ok += 1
        print(f"  OK   {row.Номер} {row.Месяц}")
    except Exception as e:
        err = str(e)
        if hasattr(e, 'excepinfo') and e.excepinfo:
            err = e.excepinfo[2] or err
        results.append({"номер": str(row.Номер), "статус": "FAIL", "ошибка": err})
        fail += 1
        print(f"  FAIL {row.Номер}: {err}")

print(f"\n=== ИТОГ: OK={ok}, FAIL={fail} ===")

log_dir = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "data", "json"))
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "20_reprovesti_finrez_pl_log.json")
with open(log_path, "w", encoding="utf-8") as f:
    json.dump({
        "ran_at": datetime.datetime.now().isoformat(),
        "total": tz.Количество(), "ok": ok, "fail": fail, "items": results,
    }, f, ensure_ascii=False, indent=2)
print(f"Log: {log_path}")
