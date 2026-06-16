# -*- coding: utf-8 -*-
"""
PILOT ПоступлениеБезналичныхДенежныхСредств "Поступление ДС с другого счёта" Шаг 01 — Baseline ДО доработки.

Документ: ПоступлениеБезналичныхДенежныхСредств Номер "00DL-006964" от 01.12.2025
  ХозОп = ПоступлениеДенежныхСредствСДругогоСчета, СуммаДокумента 9 641 228,35 UAH
  БанковскийСчет (получатель), БанковскийСчетОтправитель
  А_ПодразделениеОтправитель = "Строительство" (заполнен вручную)
  А_СписаниеБезналичныхДенежныхСредств = парное Списание Номер "000Ц-000090"

Ожидаем ДО фикса:
  - РНДС.ДенежныеСредстваВПути — строка Расход 9641228.35, Подразделение = шапки (НЕ Отправитель).
  - ПрочиеАктивыПассивы — строка Статья/Источник "Денежные средства в пути",
      Подразделение = шапки.
  - ДенежныеСредстваБезналичные — Приход 9641228.35.

Артефакт: _artifacts/pilot_postbeznal_priem_01_baseline.json
Rule #-1: запрос прогнан на реальной базе ПЕРЕД любым переносом в BSL.
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
DOC_NUM = "00DL-006964"

# --- Найти документ ---
q = erp.NewObject("Запрос")
q.УстановитьПараметр("Ном", DOC_NUM)
q.Текст = 'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.ПоступлениеБезналичныхДенежныхСредств ГДЕ Номер = &Ном'
sel = q.Выполнить().Выбрать()
if not sel.Следующий():
    print(f"[FAIL] ПоступлениеБезналичныхДенежныхСредств {DOC_NUM} не найден"); sys.exit(1)
DOC = sel.Ссылка
print(f"Документ: {S(DOC)}")


def conv(v):
    """Привести значение реквизита/поля к JSON-сериализуемому виду."""
    if v is None:
        return None
    if isinstance(v, (str, int, float, bool)):
        return v
    try:
        return str(S(v)) if erp.ЗначениеЗаполнено(v) else None
    except Exception:
        return "<obj>"


# --- Шапка ---
obj = DOC.ПолучитьОбъект()
print("\n=== Реквизиты шапки ===")
hdr = {}
for fld in ("Дата", "Номер", "Проведен", "ПометкаУдаления", "Организация", "Подразделение",
            "БанковскийСчет", "БанковскийСчетОтправитель", "ХозяйственнаяОперация",
            "А_ПодразделениеОтправитель", "А_СписаниеБезналичныхДенежныхСредств",
            "Валюта", "СуммаДокумента",
            "А_ОбработанКазна", "А_ВведенВЕРП", "А_Обработан"):
    try:
        v = getattr(obj, fld)
        sv = conv(v)
        hdr[fld] = sv
        print(f"  {fld:<42}: {'(пусто)' if sv is None else sv}")
    except Exception:
        pass

# --- Регистры ---
REGISTRY = ["ДенежныеСредстваВПути", "ПрочиеАктивыПассивы", "ДенежныеСредстваБезналичные", "ДвиженияДенежныхСредств"]
# поля-кандидаты для "Статья": Статья (ПАП) / СтатьяДвиженияДенежныхСредств (ДС-регистры)
STAT_FIELDS = ["Статья", "СтатьяДвиженияДенежныхСредств"]

snapshot = {"document": str(S(DOC)), "doc_number": DOC_NUM, "header": hdr, "movements": {}}

print("\n=== Дамп движений по регистрам ===")
for reg_name in REGISTRY:
    qq = erp.NewObject("Запрос")
    qq.УстановитьПараметр("Док", DOC)
    qq.Текст = f"ВЫБРАТЬ Р.* ИЗ РегистрНакопления.{reg_name} КАК Р ГДЕ Р.Регистратор = &Док"
    try:
        rr = qq.Выполнить().Выгрузить()
    except Exception as e:
        info = getattr(e, "excepinfo", None)
        print(f"\n  [{reg_name}] недоступен: {info[2] if info else e}"[:160])
        continue

    cols = [c.Имя for c in rr.Колонки]
    has_vd = "ВидДвижения" in cols
    has_ist = "Источник" in cols
    stat_field = next((sf for sf in STAT_FIELDS if sf in cols), None)
    print(f"\n--- {reg_name}: {rr.Количество()} строк "
          f"(ВидДвижения={'да' if has_vd else 'НЕТ'}, "
          f"Статья={stat_field or 'НЕТ'}, Источник={'да' if has_ist else 'НЕТ'}) ---")

    rows = []
    # агрегация Σ Сумма по (ВидДвижения, Подразделение, Статья, Источник)
    agg = {}
    for i in range(rr.Количество()):
        rec = rr.Получить(i)
        row = {c: conv(getattr(rec, c, None)) for c in cols}
        rows.append(row)

        vd = row.get("ВидДвижения") or "(нет)"
        sum_ = row.get("Сумма")
        if sum_ is None:
            sum_ = row.get("СуммаУпр") or row.get("СуммаВВалюте") or 0
        podr = row.get("Подразделение")
        stat = row.get(stat_field) if stat_field else None
        ist = row.get("Источник")

        key = (str(vd), str(podr), str(stat), str(ist))
        agg[key] = agg.get(key, 0) + (sum_ or 0)

    # печать агрегата
    print("  ВидДвиж     |        Σ Сумма | Подразделение            | Статья / Источник")
    agg_out = []
    for (vd, podr, stat, ist), s in sorted(agg.items()):
        podr_s = "(пусто)" if podr in ("None", None) else podr
        stat_s = "" if stat in ("None", None) else stat
        ist_s = "" if ist in ("None", None) else ist
        si = stat_s + (f"  | Источник={ist_s}" if ist_s else "")
        print(f"  {vd:<11} | {s:>14} | {str(podr_s)[:24]:<24} | {si}")
        agg_out.append({"ВидДвижения": vd, "Подразделение": podr_s,
                        "Статья": stat_s or None, "Источник": ist_s or None, "СуммаΣ": s})

    snapshot["movements"][reg_name] = {
        "count": rr.Количество(),
        "columns": cols,
        "stat_field": stat_field,
        "has_istochnik": has_ist,
        "aggregate": agg_out,
        "rows": rows,
    }

# --- ЯВНОЕ подтверждение ключевых полей baseline ---
print("\n=== ПОДТВЕРЖДЕНИЕ baseline (ДО фикса) ===")


def find_agg(reg, predicate):
    data = snapshot["movements"].get(reg)
    if not data:
        return []
    return [a for a in data["aggregate"] if predicate(a)]

# 1) ВПути.Расход — Подразделение
vputi_rashod = find_agg("ДенежныеСредстваВПути", lambda a: a["ВидДвижения"] == "Расход")
if vputi_rashod:
    for a in vputi_rashod:
        print(f"  [ВПути.Расход] Σ={a['СуммаΣ']}  Подразделение = {a['Подразделение']}")
    snapshot["confirm_vputi_rashod_podrazdelenie"] = vputi_rashod[0]["Подразделение"]
else:
    print("  [ВПути.Расход] строк НЕ найдено!")
    snapshot["confirm_vputi_rashod_podrazdelenie"] = None

# 2) ПАП-строка "Денежные средства в пути"
pap_vputi = find_agg(
    "ПрочиеАктивыПассивы",
    lambda a: (a.get("Источник") and "пути" in a["Источник"].lower())
    or (a.get("Статья") and "пути" in a["Статья"].lower()),
)
if pap_vputi:
    for a in pap_vputi:
        print(f"  [ПАП 'Денежные средства в пути'] Σ={a['СуммаΣ']}  "
              f"Подразделение = {a['Подразделение']}  Статья={a['Статья']}  Источник={a['Источник']}")
    snapshot["confirm_pap_vputi_podrazdelenie"] = pap_vputi[0]["Подразделение"]
else:
    print("  [ПАП 'Денежные средства в пути'] строк НЕ найдено!")
    snapshot["confirm_pap_vputi_podrazdelenie"] = None

# 3) Безналичные.Приход
beznal_prihod = find_agg("ДенежныеСредстваБезналичные", lambda a: a["ВидДвижения"] == "Приход")
for a in beznal_prihod:
    print(f"  [Безналичные.Приход] Σ={a['СуммаΣ']}  Подразделение = {a['Подразделение']}")

# --- Сохранить JSON ---
out = os.path.join(ART, "pilot_postbeznal_priem_01_baseline.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(snapshot, f, ensure_ascii=False, indent=2, default=str)
print(f"\n[OK] Snapshot: {out}")
