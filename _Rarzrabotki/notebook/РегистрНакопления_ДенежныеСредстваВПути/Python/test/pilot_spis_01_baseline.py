# -*- coding: utf-8 -*-
"""PILOT СписаниеБезнал Шаг 01 — Baseline 00000019546 ДО доработки."""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts"); os.makedirs(ART, exist_ok=True)

q = erp.NewObject("Запрос")
q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.СписаниеБезналичныхДенежныхСредств ГДЕ Номер = "00000019546"'
sel = q.Выполнить().Выбрать()
if not sel.Следующий(): print("[FAIL] не найден"); sys.exit(1)
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")
obj = DOC.ПолучитьОбъект()
print("\n=== Реквизиты шапки ===")
hdr = {}
for fld in ("Дата","Номер","Проведен","Организация","Подразделение","БанковскийСчет","БанковскийСчетПолучатель",
            "КассаПолучатель","ХозяйственнаяОперация","Валюта","СуммаДокумента",
            "А_ОбработанКазна","А_ВведенВЕРП","Партнер","Контрагент"):
    try:
        v = getattr(obj, fld)
        if v is None: sv = None
        elif isinstance(v,(str,int,float,bool)): sv = v
        else: sv = str(S(v)) if erp.ЗначениеЗаполнено(v) else "(пусто)"
        hdr[fld] = sv; print(f"  {fld:<25}: {sv}")
    except: pass

REGISTRY = ["ДенежныеСредстваВПути","ДенежныеСредстваБезналичные","ПрочиеАктивыПассивы","ДвиженияДенежныхСредств",
            "РасчетыСКлиентамиПоСрокам","РасчетыСПоставщикамиПоСрокам","РасчетыСКлиентами","РасчетыСПоставщиками"]
snapshot = {"document": str(S(DOC)), "header": hdr, "movements": {}}
print("\n=== Дамп движений ===")
for reg in REGISTRY:
    qq = erp.NewObject("Запрос"); qq.УстановитьПараметр("Док", DOC)
    qq.Текст = f"ВЫБРАТЬ Р.* ИЗ РегистрНакопления.{reg} КАК Р ГДЕ Р.Регистратор = &Док"
    try: rr = qq.Выполнить().Выгрузить()
    except: continue
    cols = [c.Имя for c in rr.Колонки]
    if rr.Количество() == 0: continue
    print(f"\n--- {reg}: {rr.Количество()} строк ---")
    rows = []
    for i in range(rr.Количество()):
        rec = rr.Получить(i); row = {}
        for c in cols:
            v = getattr(rec, c, None)
            if v is None: row[c] = None
            elif isinstance(v,(str,int,float,bool)): row[c] = v
            else:
                try: row[c] = str(S(v)) if erp.ЗначениеЗаполнено(v) else None
                except: row[c] = "<obj>"
        rows.append(row)
        vd = row.get("ВидДвижения","")
        sum_ = row.get("Сумма") or 0
        podr = row.get("Подразделение") or "(нет/пусто)"
        getr = row.get("Получатель") or row.get("БанковскийСчет") or ""
        otpr = row.get("Отправитель") or ""
        target = f"{otpr}→{getr}"[:55] if otpr or getr else (row.get("БанковскийСчет") or "")[:55]
        print(f"  {str(vd):<8} Σ={sum_:>14}  Подр={str(podr)[:23]:<25} {target}")
    snapshot["movements"][reg] = {"count": rr.Количество(), "rows": rows}

out = os.path.join(ART, "pilot_spis_01_baseline.json")
with open(out,"w",encoding="utf-8") as f: json.dump(snapshot,f,ensure_ascii=False,indent=2,default=str)
print(f"\n[OK] Snapshot: {out}")
