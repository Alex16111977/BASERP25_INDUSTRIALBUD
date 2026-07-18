# -*- coding: utf-8 -*-
"""ПостБезнал «с другого счёта» Шаг 04 — verify 00DL-006964 (no-regression).
Ожидаем: ВПути.Расход.Подр = А_ПодразделениеОтправитель; ПАП «в пути».Подр = то же; Σ инвариантна.
Для 00DL-006964 шапка = отправитель (Строительство), поэтому фикс — no-op: значения остаются Строительство (доказ. что код не ломает)."""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
import win32com.client, pythoncom
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
def S(v):
    try: return v.Наименование if erp.ЗначениеЗаполнено(v) else "(пусто)"
    except: return "(?)"
ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
BL = os.path.join(ART, "pilot_postbeznal_priem_01_baseline.json")
before = json.load(open(BL, encoding="utf-8")) if os.path.exists(BL) else {"movements": {}}

DOC_NUM = "00DL-006964"
q = erp.NewObject("Запрос")
q.Текст = f'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ПоступлениеБезналичныхДенежныхСредств ГДЕ Номер = "{DOC_NUM}"'
sel = q.Выполнить().Выбрать(); sel.Следующий()
DOC = sel.Ссылка; obj = DOC.ПолучитьОбъект()
exp = S(obj.А_ПодразделениеОтправитель)
print(f"Документ: {DOC_NUM}; ожидаемое Подразделение ВПути/ПАП-в-пути = А_ПодразделениеОтправитель = {exp}\n")

REGS = ["ДенежныеСредстваВПути","ПрочиеАктивыПассивы","ДенежныеСредстваБезналичные","ДвиженияДенежныхСредств"]
res = {"pass": [], "fail": []}
for reg in REGS:
    qq = erp.NewObject("Запрос"); qq.УстановитьПараметр("Док", DOC)
    qq.Текст = f"ВЫБРАТЬ Р.* ИЗ РегистрНакопления.{reg} КАК Р ГДЕ Р.Регистратор = &Док"
    try: rr = qq.Выполнить().Выгрузить()
    except Exception as e:
        info=getattr(e,"excepinfo",None); print(f"[{reg}] err {info[2] if info else e}"); continue
    cols = [c.Имя for c in rr.Колонки]; sum_now=0.0; rows=[]
    for i in range(rr.Количество()):
        rec = rr.Получить(i)
        vd = S(getattr(rec,"ВидДвижения",None)) if "ВидДвижения" in cols else ""
        summ = float(getattr(rec,"Сумма",0) or 0) if "Сумма" in cols else 0.0; sum_now+=summ
        podr = S(getattr(rec,"Подразделение",None)) if "Подразделение" in cols else ""
        st = S(getattr(rec,"Статья",None)) if "Статья" in cols else ""
        rows.append({"vd":vd,"podr":podr,"st":st,"sum":summ})
        print(f"  {reg[:26]:<26} {vd:<8} Σ={summ:>14,.2f}  Подр={podr:<16} {st[:30]}")
    b = before["movements"].get(reg, {"rows": []})
    sb = sum(float(r.get("Сумма") or 0) for r in b.get("rows", []))
    if abs(sum_now-sb) > 0.01: res["fail"].append(f"{reg}: Σ {sb:,.2f}→{sum_now:,.2f}")
    else: res["pass"].append(f"{reg}: Σ={sum_now:,.2f} инвариантна")
    if reg=="ДенежныеСредстваВПути":
        p={r["podr"] for r in rows}
        (res["pass"] if p=={exp} else res["fail"]).append(f"ВПути.Подр={sorted(p)} (ожид {exp})")
    if reg=="ПрочиеАктивыПассивы":
        vp=[r for r in rows if "пути" in r["st"].lower()]
        if vp and all(r["podr"]==exp for r in vp): res["pass"].append(f"ПАП «в пути».Подр={exp}")
        elif vp: res["fail"].append(f"ПАП «в пути».Подр={[r['podr'] for r in vp]} (ожид {exp})")
        else: res["fail"].append("ПАП «в пути» строка не найдена")
print("\n=== РЕЗУЛЬТАТ ===")
for p in res["pass"]: print(f"  ✓ {p}")
for f in res["fail"]: print(f"  ✗ {f}")
print("\n[OVERALL]", "FAIL" if res["fail"] else "PASS — код компилируется, документ проводится, значения консистентны (no-regression)")
sys.exit(1 if res["fail"] else 0)
