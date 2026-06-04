# -*- coding: utf-8 -*-
"""
ПИЛОТ Шаг 01 — Baseline ПКО 000Ц-000001 ДО доработки.

Дамп всех движений документа во всех регистрах + реквизиты шапки + Касса.Подразделение.

Образец: _Rarzrabotki/Python/test/train_repost_single_doc.py (дамп ДО/ПОСЛЕ + сравнение).

Сохраняет:
  _artifacts/pilot_01_baseline.json — снимок движений (используется в pilot_04 для diff)
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

# --- Найти ПКО 000Ц-000001 ---
q = erp.NewObject("Запрос")
q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ПриходныйКассовыйОрдер ГДЕ Номер = "000Ц-000001"'
sel = q.Выполнить().Выбрать()
if not sel.Следующий():
    print("[FAIL] ПКО 000Ц-000001 не найден"); sys.exit(1)
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")

# --- Реквизиты шапки ---
obj = DOC.ПолучитьОбъект()
print("\n=== Реквизиты шапки ===")
hdr = {}
for fld in ("Дата", "Номер", "Проведен", "ПометкаУдаления", "Организация", "Подразделение",
            "Касса", "ХозяйственнаяОперация", "Валюта", "ВалютаКонвертации", "СуммаДокумента",
            "А_ОбработанКазна", "А_ВведенВЕРП", "А_Обработан"):
    try:
        v = getattr(obj, fld)
        if v is None:
            sv = None
        elif isinstance(v, (str, int, float, bool)):
            sv = v
        else:
            sv = str(S(v)) if erp.ЗначениеЗаполнено(v) else "(пусто)"
        hdr[fld] = sv
        print(f"  {fld:<25}: {sv}")
    except Exception as e:
        hdr[fld] = f"<err: {e}>"

# --- Подразделение из Кассы (target) ---
podr_kassa = None
try:
    if erp.ЗначениеЗаполнено(obj.Касса):
        podr_kassa = obj.Касса.Подразделение
        print(f"\n  Касса.Подразделение (целевое для движения): "
              f"{S(podr_kassa) if erp.ЗначениеЗаполнено(podr_kassa) else '(пусто)'}")
except: pass

# --- Дамп движений по регистрам которые пишет ПКО ---
REGISTRY = [
    "ДенежныеСредстваВПути",
    "ДенежныеСредстваНаличные",
    "ПрочиеАктивыПассивы",
    "ДвиженияДенежныхСредств",
]

snapshot = {"document": str(S(DOC)), "header": hdr,
            "kassa_podrazdelenie": str(S(podr_kassa)) if podr_kassa and erp.ЗначениеЗаполнено(podr_kassa) else None,
            "movements": {}}

print("\n=== Дамп движений по регистрам ===")
for reg_name in REGISTRY:
    qq = erp.NewObject("Запрос")
    qq.УстановитьПараметр("Док", DOC)
    # Динамический SELECT * через метаданные не поддерживается — берём ключевые поля
    qq.Текст = f"""
    ВЫБРАТЬ Р.* ИЗ РегистрНакопления.{reg_name} КАК Р ГДЕ Р.Регистратор = &Док
    """
    try:
        rr = qq.Выполнить().Выгрузить()
    except Exception as e:
        info = getattr(e, "excepinfo", None)
        print(f"\n  [{reg_name}] не доступен: {info[2] if info else e}"[:120])
        continue

    cols = [c.Имя for c in rr.Колонки]
    print(f"\n--- {reg_name}: {rr.Количество()} строк ---")
    rows = []
    # Кратко выводим: ВидДвижения + ключевые поля + Сумма
    for i in range(rr.Количество()):
        rec = rr.Получить(i)
        row = {}
        for c in cols:
            v = getattr(rec, c, None)
            if v is None:
                row[c] = None
            elif isinstance(v, (str, int, float, bool)):
                row[c] = v
            else:
                try: row[c] = str(S(v)) if erp.ЗначениеЗаполнено(v) else None
                except: row[c] = "<obj>"
        rows.append(row)
        # Печать ключевых полей
        vd = row.get("ВидДвижения", "")
        sum_ = row.get("Сумма") or row.get("СуммаУпр") or 0
        podr = row.get("Подразделение") or "(нет поля или пусто)"
        # для РНДС.ВПути
        getr = row.get("Получатель") or ""
        otpr = row.get("Отправитель") or ""
        bs = row.get("БанковскийСчет") or ""
        kassa = row.get("Касса") or ""
        target = bs or kassa or f"{otpr}→{getr}"
        print(f"  {str(vd):<10} Σ={sum_:>14}  Подр={str(podr)[:25]:<25}  {str(target)[:60]}")
    snapshot["movements"][reg_name] = {"count": rr.Количество(), "rows": rows}

# Сохранить snapshot
out = os.path.join(ART, "pilot_01_baseline.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(snapshot, f, ensure_ascii=False, indent=2, default=str)
print(f"\n[OK] Snapshot: {out}")
print(f"     Используется pilot_04_verify.py для сравнения ДО vs ПОСЛЕ")
