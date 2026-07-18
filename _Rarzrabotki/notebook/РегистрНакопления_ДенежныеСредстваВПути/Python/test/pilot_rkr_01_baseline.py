# -*- coding: utf-8 -*-
"""PILOT РасчетКурсовых Шаг 01 — Baseline 000Ц-000007 (январь 2026)."""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String
ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts"); os.makedirs(ART, exist_ok=True)

q = erp.NewObject("Запрос")
q.УстановитьПараметр("ХО", erp.Перечисления.ХозяйственныеОперации.ПереоценкаДенежныхСредств)
q.Текст = ('ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.РасчетКурсовыхРазниц '
           'ГДЕ Номер = "000Ц-000007" И ХозяйственнаяОперация = &ХО И ГОД(Дата) = 2026')
sel = q.Выполнить().Выбрать()
if not sel.Следующий(): print("[FAIL]"); sys.exit(1)
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")
obj = DOC.ПолучитьОбъект()
hdr = {}
for fld in ("Дата","Номер","Проведен","Организация","ХозяйственнаяОперация","Подразделение"):
    try:
        v = getattr(obj, fld, None)
        if v is None: sv = None
        elif isinstance(v,(str,int,float,bool)): sv = v
        else: sv = str(S(v)) if erp.ЗначениеЗаполнено(v) else "(пусто)"
        hdr[fld] = sv; print(f"  {fld:<22}: {sv}")
    except: pass

REGISTRY = ["ДенежныеСредстваВПути"]
snapshot = {"document": str(S(DOC)), "header": hdr, "movements": {}}
print("\n=== Движения в РНДС.ВПути ===")
for reg in REGISTRY:
    qq = erp.NewObject("Запрос"); qq.УстановитьПараметр("Док", DOC)
    qq.Текст = f"ВЫБРАТЬ Р.* ИЗ РегистрНакопления.{reg} КАК Р ГДЕ Р.Регистратор = &Док"
    rr = qq.Выполнить().Выгрузить()
    cols = [c.Имя for c in rr.Колонки]
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
        sup = row.get("СуммаУпр") or 0
        srg = row.get("СуммаРегл") or 0
        podr = row.get("Подразделение") or "(нет/пусто)"
        val = row.get("Валюта") or ""
        get_ = row.get("Получатель") or ""
        otp_ = row.get("Отправитель") or ""
        print(f"  {str(vd):<8} {val:<5} Σ={sum_:>6} Упр={sup:>10} Регл={srg:>10}  Подр={str(podr)[:20]:<22} {otp_} → {get_}")
    snapshot["movements"][reg] = {"count": rr.Количество(), "rows": rows}
out = os.path.join(ART, "pilot_rkr_01_baseline.json")
with open(out,"w",encoding="utf-8") as f: json.dump(snapshot,f,ensure_ascii=False,indent=2,default=str)
print(f"\n[OK] Snapshot: {out}")
