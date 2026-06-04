# -*- coding: utf-8 -*-
"""
PILOT РКО Шаг 01 — Baseline РасходныйКассовыйОрдер N0000053020 ДО доработки.

Аналог pilot_01_baseline.py но для РКО.
Артефакт: _artifacts/pilot_rko_01_baseline.json
"""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
os.makedirs(ART, exist_ok=True)

q = erp.NewObject("Запрос")
q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.РасходныйКассовыйОрдер ГДЕ Номер = "N0000053020"'
sel = q.Выполнить().Выбрать()
if not sel.Следующий():
    print("[FAIL] РКО N0000053020 не найден"); sys.exit(1)
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")

obj = DOC.ПолучитьОбъект()
print("\n=== Реквизиты шапки ===")
hdr = {}
for fld in ("Дата", "Номер", "Проведен", "ПометкаУдаления", "Организация", "Подразделение",
            "Касса", "КассаПолучатель", "БанковскийСчет", "ХозяйственнаяОперация",
            "Валюта", "ВалютаКонвертации", "СуммаДокумента",
            "А_ОбработанКазна", "А_ВведенВЕРП", "А_Обработан"):
    try:
        v = getattr(obj, fld)
        if v is None: sv = None
        elif isinstance(v, (str, int, float, bool)): sv = v
        else: sv = str(S(v)) if erp.ЗначениеЗаполнено(v) else "(пусто)"
        hdr[fld] = sv
        print(f"  {fld:<25}: {sv}")
    except: pass

REGISTRY = ["ДенежныеСредстваВПути", "ДенежныеСредстваНаличные", "ПрочиеАктивыПассивы", "ДвиженияДенежныхСредств"]
snapshot = {"document": str(S(DOC)), "header": hdr, "movements": {}}

print("\n=== Дамп движений ===")
for reg_name in REGISTRY:
    qq = erp.NewObject("Запрос")
    qq.УстановитьПараметр("Док", DOC)
    qq.Текст = f"ВЫБРАТЬ Р.* ИЗ РегистрНакопления.{reg_name} КАК Р ГДЕ Р.Регистратор = &Док"
    try: rr = qq.Выполнить().Выгрузить()
    except Exception as e:
        info = getattr(e, "excepinfo", None)
        print(f"\n  [{reg_name}] недоступен: {info[2] if info else e}"[:120]); continue
    cols = [c.Имя for c in rr.Колонки]
    print(f"\n--- {reg_name}: {rr.Количество()} строк ---")
    rows = []
    for i in range(rr.Количество()):
        rec = rr.Получить(i)
        row = {}
        for c in cols:
            v = getattr(rec, c, None)
            if v is None: row[c] = None
            elif isinstance(v, (str, int, float, bool)): row[c] = v
            else:
                try: row[c] = str(S(v)) if erp.ЗначениеЗаполнено(v) else None
                except: row[c] = "<obj>"
        rows.append(row)
        vd = row.get("ВидДвижения", "")
        sum_ = row.get("Сумма") or row.get("СуммаУпр") or 0
        podr = row.get("Подразделение") or "(нет поля или пусто)"
        getr = row.get("Получатель") or ""
        otpr = row.get("Отправитель") or ""
        bs = row.get("БанковскийСчет") or ""
        kassa = row.get("Касса") or ""
        target = bs or kassa or f"{otpr}→{getr}"
        print(f"  {str(vd):<10} Σ={sum_:>14}  Подр={str(podr)[:25]:<25}  {str(target)[:60]}")
    snapshot["movements"][reg_name] = {"count": rr.Количество(), "rows": rows}

out = os.path.join(ART, "pilot_rko_01_baseline.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(snapshot, f, ensure_ascii=False, indent=2, default=str)
print(f"\n[OK] Snapshot: {out}")
