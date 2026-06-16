# -*- coding: utf-8 -*-
# Прототип ДОДЕЛКИ движка: чистит остаток 171 ссылки.
# A. Удаляет мои помеченные дубль-ключи (КлючиАналитикиУчетаПоПартнерам, Партнер=Неизв, ПометкаУдаления).
# B. РеестрДокументов: Партнер(измерение)=Неизв -> target от контрагента регистратора -> удалить+записать.
# C. Платёжные документы: Партнер в шапке И/ИЛИ в РасшифровкаПлатежа -> обойти, target от шапки Контрагент.
# Резолвер берём из базы: А_Привилегированный.ОбеспечитьРеальногоПартнераПоКонтрагенту.
import sys
sys.path.insert(0, r'C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test')
import _common_neizv as c

erp = c.connect()
P = c.neizv(erp)


def resolve(контр):
    if not erp.ЗначениеЗаполнено(контр):
        return None
    # резолвер требует именно Контрагенты; Организации/прочее кинут ошибку -> ловим
    try:
        return erp.А_Привилегированный.ОбеспечитьРеальногоПартнераПоКонтрагенту(контр, None, True)
    except Exception:
        return None


def refs_count(ref):
    arr = erp.NewObject("Массив"); arr.Добавить(ref)
    return erp.НайтиПоСсылкам(arr).Количество()


# ---------- A: удалить помеченные дубль-ключи ----------
print("=== A: удаление помеченных дубль-ключей ===")
q = erp.NewObject("Запрос")
q.Text = ("ВЫБРАТЬ Ссылка ИЗ Справочник.КлючиАналитикиУчетаПоПартнерам "
          "ГДЕ Партнер = &P И ПометкаУдаления")
q.УстановитьПараметр("P", P)
s = q.Execute().Выбрать()
keys = []
while s.Следующий():
    keys.append(s.Ссылка)
deleted = skipped = 0
for k in keys:
    if refs_count(k) == 0:
        k.ПолучитьОбъект().Удалить()
        deleted += 1
    else:
        skipped += 1
print(f"  удалено сирот-ключей: {deleted}, пропущено (есть ссылки): {skipped}")


# ---------- B: РеестрДокументов ----------
print("\n=== B: РеестрДокументов (Партнер=измерение -> удалить+записать) ===")
RES_FIELDS = ["НомерДокументаИБ", "Статус", "Ответственный", "ДополнительнаяЗапись", "Дополнительно",
              "Комментарий", "Проведен", "ПометкаУдаления", "ДатаПервичногоДокумента",
              "НомерПервичногоДокумента", "Сумма", "Валюта", "Договор", "НаправлениеДеятельности",
              "ДатаОтраженияВУчете"]
DIM_FIELDS = ["ТипСсылки", "ХозяйственнаяОперация", "Организация", "МестоХранения", "Контрагент",
              "Подразделение", "ДатаДокументаИБ", "Ссылка", "РазделительЗаписи"]
q = erp.NewObject("Запрос")
q.Text = ("ВЫБРАТЬ Партнер, " + ",".join(DIM_FIELDS + RES_FIELDS) +
          " ИЗ РегистрСведений.РеестрДокументов ГДЕ Партнер = &P")
q.УстановитьПараметр("P", P)
s = q.Execute().Выбрать()
rows = []
while s.Следующий():
    rec = {}
    for f in DIM_FIELDS + RES_FIELDS:
        rec[f] = getattr(s, f)
    rec["Ссылка_рег"] = s.Ссылка  # регистратор-документ
    rows.append(rec)
print(f"  записей РеестрДокументов с Неизв: {len(rows)}")
fixed_rd = skip_rd = 0
for rec in rows:
    контр = None
    try:
        контр = rec["Ссылка_рег"].Контрагент
    except Exception:
        контр = None
    target = resolve(контр) if (контр is not None and erp.ЗначениеЗаполнено(контр)) else None
    if target is None or c.same(erp, target, P):
        skip_rd += 1
        continue
    нов = erp.РегистрыСведений.РеестрДокументов.СоздатьМенеджерЗаписи()
    нов.Партнер = target
    for f in DIM_FIELDS:
        setattr(нов, f, rec[f])
    for f in RES_FIELDS:
        setattr(нов, f, rec[f])
    нов.Записать()
    стар = erp.РегистрыСведений.РеестрДокументов.СоздатьМенеджерЗаписи()
    стар.Партнер = P
    for f in DIM_FIELDS:
        setattr(стар, f, rec[f])
    стар.Удалить()
    fixed_rd += 1
print(f"  исправлено: {fixed_rd}, пропущено (нет контрагента/партнёра): {skip_rd}")


# ---------- C: платёжные документы (шапка + РасшифровкаПлатежа) ----------
print("\n=== C: платёжные документы (шапка + ТЧ РасшифровкаПлатежа) ===")
PAY_DOCS = ["ЗаявкаНаРасходованиеДенежныхСредств", "СписаниеБезналичныхДенежныхСредств",
            "ПоступлениеБезналичныхДенежныхСредств", "ПриобретениеТоваровУслуг"]
ТипП = erp.NewObject("Массив")  # not used; helper below
for dt in PAY_DOCS:
    # документы с Неизв в шапке ИЛИ в РасшифровкаПлатежа
    q = erp.NewObject("Запрос")
    q.Text = (
        "ВЫБРАТЬ РАЗЛИЧНЫЕ Ссылка КАК С ИЗ Документ." + dt + " ГДЕ Партнер = &P "
        "ОБЪЕДИНИТЬ "
        "ВЫБРАТЬ РАЗЛИЧНЫЕ Ссылка ИЗ Документ." + dt + ".РасшифровкаПлатежа ГДЕ Партнер = &P")
    q.УстановитьПараметр("P", P)
    try:
        s = q.Execute().Выбрать()
    except Exception as e:
        info = getattr(e, 'excepinfo', None)
        print(f"  {dt}: запрос FAIL {info[2] if info else e}")
        continue
    docs = []
    while s.Следующий():
        docs.append(s.С)
    cnt = 0
    for dref in docs:
        obj = dref.ПолучитьОбъект()
        контр = None
        try:
            контр = obj.Контрагент
        except Exception:
            контр = None
        target = resolve(контр) if (контр is not None and erp.ЗначениеЗаполнено(контр)) else None
        if target is None or c.same(erp, target, P):
            continue
        changed = False
        # шапка Партнер
        try:
            if c.same(erp, obj.Партнер, P):
                obj.Партнер = target; changed = True
        except Exception:
            pass
        # ТЧ РасшифровкаПлатежа
        try:
            for row in obj.РасшифровкаПлатежа:
                if c.same(erp, row.Партнер, P):
                    row.Партнер = target; changed = True
        except Exception:
            pass
        if changed:
            obj.ОбменДанными.Загрузка = True
            obj.Записать()
            cnt += 1
    print(f"  {dt}: исправлено документов {cnt}")


# ---------- ИТОГ ----------
total, agg = c.count_refs(erp, P)
print("\n=== ОСТАТОК ссылок на Неизвестного ===", total)
for k in sorted(agg, key=lambda x: -agg[x]):
    print(f"  {agg[k]:5}  {k}")
