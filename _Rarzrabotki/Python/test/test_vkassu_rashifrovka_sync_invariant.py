"""
Регресс-гейт для фикса синхронизации подформы расшифровки ВКассу.

ВАЖНО: ТЧ Состав хранит только {ИдентификаторСтроки, ФизическоеЛицо}. Колонка
«К выплате» (КВыплате) — это AdditionalColumns атрибута Объект.Состав НА ФОРМЕ
(пересчитывается ПриПолученииДанныхСтрокиСостава из ТЧ Зарплата) и через сырой
COM-объект документа НЕ читается. Поэтому инвариант проверяется по реальным
хранимым ТЧ (Зарплата, А_Расшифровка) и реквизиту СуммаПоДокументу.

Проверяет:
1. ТЧ Зарплата имеет все реквизиты, на которые опирается фикс
   (ЗаполнитьЗначенияСвойств из строки расшифровки + Итог по КВыплате).
2. ТЧ Состав имеет ИдентификаторСтроки / ФизическоеЛицо.
3. Реквизит шапки СуммаПоДокументу существует.
4. Инвариант данных на документе 000Ц-000005:
   - для каждого ИдентификаторСтроки:  Σ Зарплата.КВыплате == Σ А_Расшифровка.КВыплате
   - по документу:                     Σ Зарплата.КВыплате == СуммаПоДокументу
   - нет orphan: каждый ИдентификаторСтроки Зарплаты/Расшифровки есть в Составе
"""
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)

EPS = 0.01
fails = []

def uid(ref):
    """Строковый ключ ссылки/UUID через серверный XMLСтрока (надёжно для COM)."""
    try:
        return erp.XMLСтрока(ref)
    except Exception:
        return str(ref)

# --- 1/2. Метаданные ТЧ ---------------------------------------------------
meta_doc = erp.Метаданные.Документы.Найти("ВедомостьНаВыплатуЗарплатыВКассу")
if meta_doc is None:
    print("FAIL: документ не найден"); sys.exit(1)

def tch_attrs(name):
    for i in range(meta_doc.ТабличныеЧасти.Количество()):
        t = meta_doc.ТабличныеЧасти.Получить(i)
        if str(t.Имя) == name:
            return [str(t.Реквизиты.Получить(j).Имя) for j in range(t.Реквизиты.Количество())]
    return None

zp_attrs = tch_attrs("Зарплата")
sostav_attrs = tch_attrs("Состав")

req_zp = ["ИдентификаторСтроки", "Сотрудник", "ФизическоеЛицо", "Подразделение",
          "ПериодВзаиморасчетов", "СтатьяФинансирования", "СтатьяРасходов",
          "ДокументОснование", "КВыплате", "КомпенсацияЗаЗадержкуЗарплаты",
          "ГруппаУчетаНачислений"]
req_sostav = ["ИдентификаторСтроки", "ФизическоеЛицо"]

if zp_attrs is None:
    fails.append("ТЧ Зарплата не найдена")
else:
    miss = [a for a in req_zp if a not in zp_attrs]
    if miss: fails.append(f"ТЧ Зарплата без реквизитов: {miss}")
    else: print(f"OK: ТЧ Зарплата имеет все {len(req_zp)} нужных реквизитов")

if sostav_attrs is None:
    fails.append("ТЧ Состав не найдена")
else:
    miss = [a for a in req_sostav if a not in sostav_attrs]
    if miss: fails.append(f"ТЧ Состав без реквизитов: {miss}")
    else: print(f"OK: ТЧ Состав имеет все {len(req_sostav)} нужных реквизитов")

hdr = [str(meta_doc.Реквизиты.Получить(i).Имя) for i in range(meta_doc.Реквизиты.Количество())]
if "СуммаПоДокументу" not in hdr:
    fails.append("Реквизит шапки СуммаПоДокументу не найден")
else:
    print("OK: реквизит шапки СуммаПоДокументу существует")

# --- 4. Инвариант данных на 000Ц-000005 ----------------------------------
q = erp.NewObject("Запрос")
q.Text = ("ВЫБРАТЬ Ссылка ИЗ Документ.ВедомостьНаВыплатуЗарплатыВКассу "
          "ГДЕ Номер = &Н")
q.SetParameter("Н", "000Ц-000005")
sel = q.Execute().Выбрать()
if not sel.Следующий():
    print("WARN: документ 000Ц-000005 не найден — пропускаю проверку инварианта")
else:
    obj = sel.Ссылка.ПолучитьОбъект()

    # множество ИдентификаторСтроки из ТЧ Состав (карта ФЛ→UID)
    sostav_ids = set()
    for i in range(obj.Состав.Количество()):
        sostav_ids.add(uid(obj.Состав.Получить(i).ИдентификаторСтроки))

    # суммы Зарплата по ИдентификаторСтроки
    zp_by_id = {}
    zp_total = 0.0
    zp_orphans = set()
    for i in range(obj.Зарплата.Количество()):
        r = obj.Зарплата.Получить(i)
        k = uid(r.ИдентификаторСтроки)
        zp_by_id[k] = zp_by_id.get(k, 0.0) + float(r.КВыплате)
        zp_total += float(r.КВыплате)
        if k not in sostav_ids:
            zp_orphans.add(k)

    # суммы А_Расшифровка по ИдентификаторСтроки
    rs_by_id = {}
    rs_total = 0.0
    rs_orphans = set()
    for i in range(obj.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Количество()):
        r = obj.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Получить(i)
        k = uid(r.ИдентификаторСтроки)
        rs_by_id[k] = rs_by_id.get(k, 0.0) + float(r.КВыплате)
        rs_total += float(r.КВыплате)
        if k not in sostav_ids:
            rs_orphans.add(k)

    doc_total = float(obj.СуммаПоДокументу)
    print(f"\n000Ц-000005: ΣЗарплата={zp_total:.2f}  ΣРасшифровка={rs_total:.2f}  "
          f"СуммаПоДокументу={doc_total:.2f}  (строк Состав={len(sostav_ids)})")

    # per-ИдентификаторСтроки: Зарплата == Расшифровка
    mismatch = []
    for k in set(zp_by_id) | set(rs_by_id):
        if abs(zp_by_id.get(k, 0.0) - rs_by_id.get(k, 0.0)) > EPS:
            mismatch.append(f"{k}: Зарплата={zp_by_id.get(k,0.0)} != Расшифровка={rs_by_id.get(k,0.0)}")
    if mismatch:
        fails.append("per-UID расхождение Зарплата vs Расшифровка:\n  " + "\n  ".join(mismatch))
    else:
        print("OK: per-UID Σ Зарплата.КВыплате == Σ А_Расшифровка.КВыплате")

    if abs(zp_total - doc_total) > EPS:
        fails.append(f"ΣЗарплата({zp_total}) != СуммаПоДокументу({doc_total})")
    else:
        print("OK: Σ Зарплата.КВыплате == СуммаПоДокументу")

    if zp_orphans or rs_orphans:
        fails.append(f"orphan-строки без матча в Составе: Зарплата={zp_orphans} Расшифровка={rs_orphans}")
    else:
        print("OK: orphan-строк нет (все ИдентификаторСтроки есть в Составе)")

# --- итог -----------------------------------------------------------------
if fails:
    print("\nFAIL:")
    for f in fails: print(" -", f)
    sys.exit(1)
print("\nPASS: метаданные ТЧ корректны, инвариант на 000Ц-000005 соблюдён")
