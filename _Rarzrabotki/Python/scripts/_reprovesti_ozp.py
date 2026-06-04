# -*- coding: utf-8 -*-
"""Массовый перепровод ОЗП за 2024-01-01 .. 2026-12-31.
Лог: data/json/09_reprovesti_log.json
"""
import sys, io, json, os, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# Параметры
q = conn.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ Ссылка, Номер, Дата, Организация
ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете
ГДЕ Проведен И Дата МЕЖДУ &НМ И &КМ
УПОРЯДОЧИТЬ ПО Дата, Номер
"""
q.УстановитьПараметр("НМ", datetime.datetime(2024,1,1))
q.УстановитьПараметр("КМ", datetime.datetime(2026,12,31,23,59,59))
tz = q.Выполнить().Выгрузить()

print(f"Найдено документов ОЗП к перепроводу: {tz.Количество()}\n")

mode = conn.PredefinedValue("РежимЗаписиДокумента.Проведение")
results = []
ok_count = 0
fail_count = 0

for i in range(tz.Количество()):
    row = tz.Получить(i)
    ref = row.Ссылка
    obj = ref.ПолучитьОбъект()
    if obj is None:
        results.append({
            "doc": str(row.Номер), "date": str(row.Дата), "status": "FAIL",
            "error": "obj is None (broken ref?)"
        })
        fail_count += 1
        print(f"  FAIL {row.Номер} {row.Дата}: obj is None")
        continue
    try:
        obj.Записать(mode)
        results.append({
            "doc": str(row.Номер), "date": str(row.Дата), "org": str(row.Организация),
            "status": "OK"
        })
        ok_count += 1
        print(f"  OK   {row.Номер} {row.Дата}  {row.Организация}")
    except Exception as e:
        err = str(e)
        if hasattr(e, 'excepinfo') and e.excepinfo:
            err = e.excepinfo[2] or err
        results.append({
            "doc": str(row.Номер), "date": str(row.Дата), "status": "FAIL", "error": err
        })
        fail_count += 1
        print(f"  FAIL {row.Номер} {row.Дата}: {err}")

print(f"\n=== ИТОГ: OK={ok_count}, FAIL={fail_count} ===")

log_dir = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "data", "json"))
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "09_reprovesti_log.json")
data = {
    "ran_at": datetime.datetime.now().isoformat(),
    "total": tz.Количество(),
    "ok": ok_count,
    "fail": fail_count,
    "items": results,
}
with open(log_path, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
print(f"Log: {log_path}")
