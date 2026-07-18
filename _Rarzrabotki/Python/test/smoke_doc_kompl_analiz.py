# -*- coding: utf-8 -*-
"""Smoke 1: Документ.РасчетКомплектаций — расчёт, инварианты, таблицы списания, статус, persist.
Get-or-create по маркеру в Комментарии; документ НЕ удаляется и НЕ проводится."""
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MARKER = "SMOKE_DOC_KOMPL_v1"
PERIOD = datetime.datetime(2026, 7, 7)

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')


def get_or_create_doc():
    q = buh.NewObject("Запрос")
    q.Text = ("ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Ссылка ИЗ Документ.РасчетКомплектаций КАК Д "
              "ГДЕ ВЫРАЗИТЬ(Д.Комментарий КАК Строка(200)) = &М И НЕ Д.ПометкаУдаления")
    q.SetParameter("М", MARKER)
    r = q.Execute().Выгрузить()
    if r.Количество() > 0:
        return r.Получить(0).Ссылка.ПолучитьОбъект()
    d = buh.Документы.РасчетКомплектаций.СоздатьДокумент()
    d.Дата = datetime.datetime.now().replace(microsecond=0)
    d.Заполнить(None)
    d.Комментарий = MARKER
    return d


# --- спецификация ---
spec = buh.Справочники.СтруктураСебестоимости.НайтиПоКоду("000000004")
assert spec is not None and not spec.Пустая(), "спец 000000004 не найдена"

# --- склады с остатками (сч.20/22/28 на дату), приоритет МХП ОР ---
sch = buh.NewObject("Массив")
for kod in ("20", "22", "28"):
    sch.Добавить(buh.ПланыСчетов.Хозрасчетный.НайтиПоКоду(kod))
q = buh.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ ПЕРВЫЕ 50
	ВЫРАЗИТЬ(Ост.Субконто3 КАК Справочник.Склады) КАК Склад,
	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ Ост.Субконто1) КАК Позиций
ИЗ
	РегистрБухгалтерии.Хозрасчетный.Остатки(&НаДату, Счет В ИЕРАРХИИ (&Счета), , ) КАК Ост
ГДЕ
	Ост.КоличествоОстаток > 0

СГРУППИРОВАТЬ ПО
	ВЫРАЗИТЬ(Ост.Субконто3 КАК Справочник.Склады)

УПОРЯДОЧИТЬ ПО
	Позиций УБЫВ
"""
q.SetParameter("НаДату", datetime.datetime(2026, 7, 7, 23, 59, 59))
q.SetParameter("Счета", sch)
rows = q.Execute().Выгрузить()
assert rows.Количество() > 0, "нет складов с остатками"
sklady = []
for i in range(rows.Количество()):
    s = rows.Получить(i).Склад
    nm = s.Наименование
    sklady.append((s, nm))
mhp = [s for s, nm in sklady if "МХП" in nm.upper() or "ОРІЛЬ" in nm.upper()]
chosen = mhp[:2] if mhp else [sklady[0][0]]
print("Склады теста:", ", ".join(s.Наименование for s in chosen))

# --- документ ---
doc = get_or_create_doc()
doc.Спецификация = spec
doc.Период = PERIOD
doc.СкладыОстатков.Очистить()
for s in chosen:
    doc.СкладыОстатков.Добавить().Склад = s

doc.РассчитатьАнализ()

n = doc.ТабличнаяЧастьОстатков.Количество()
assert n > 0, "анализ пуст"
sum_ost = sum_vn = sum_pn = 0.0
rows_vn = []   # (Номенклатура строка, ВНорме, ВНормеСумма)
rows_pn = []
bad = 0
for i in range(n):
    r = doc.ТабличнаяЧастьОстатков.Получить(i)
    ost, vn, pn = float(r.Остаток), float(r.ВНорме), float(r.ПонадНорму)
    if abs(ost - (vn + pn)) > 1e-6:
        bad += 1
    sum_ost += ost; sum_vn += vn; sum_pn += pn
    if vn > 0:
        rows_vn.append((vn, float(r.ВНормеСумма)))
    if pn > 0:
        rows_pn.append((pn, float(r.ПонадНормуСумма)))
assert bad == 0, f"нарушен построчный баланс: {bad} строк"
print(f"Анализ: строк={n}, Σост={sum_ost:.3f} = ВНорме {sum_vn:.3f} + Понад {sum_pn:.3f}")

# --- таблицы списания согласованы с анализом ---
npn = doc.СписаниеПоНормам.Количество()
nsn = doc.СписаниеСверхНормы.Количество()
assert npn == len(rows_vn), f"СписаниеПоНормам: {npn} != {len(rows_vn)}"
assert nsn == len(rows_pn), f"СписаниеСверхНормы: {nsn} != {len(rows_pn)}"
sum_norm_kol = sum_norm_sum = 0.0
for i in range(npn):
    r = doc.СписаниеПоНормам.Получить(i)
    vn, vns = rows_vn[i]
    assert abs(float(r.Количество) - vn) < 1e-6, f"строка {i}: кол-во {r.Количество} != {vn}"
    assert abs(float(r.Сумма) * 1.2 - vns) <= 0.02, f"строка {i}: НДС-связка {float(r.Сумма)*1.2} vs {vns}"
    assert str(r.Счет) != "", "пустой счёт"
    sum_norm_kol += float(r.Количество); sum_norm_sum += float(r.Сумма)
sum_over_kol = 0.0
prichina_ok = True
for i in range(nsn):
    r = doc.СписаниеСверхНормы.Получить(i)
    pn, pns = rows_pn[i]
    assert abs(float(r.Количество) - pn) < 1e-6
    assert abs(float(r.Сумма) * 1.2 - pns) <= 0.02
    sum_over_kol += float(r.Количество)
assert abs(sum_norm_kol - sum_vn) < 1e-4
assert abs(sum_over_kol - sum_pn) < 1e-4
print(f"Списание: за нормами {npn} строк (Σ {sum_norm_kol:.3f}), понад {nsn} строк (Σ {sum_over_kol:.3f}); суммы без НДС ок")

# --- статус ---
st = buh.String(doc.Статус)
assert st in ("Розрахунок виконано", "Расчет выполнен"), f"статус: {st}"
print("Статус:", st)

# --- persist ---
doc.Записать()
num = doc.Номер
q2 = buh.NewObject("Запрос")
q2.Text = ("ВЫБРАТЬ Д.Ссылка КАК Ссылка ИЗ Документ.РасчетКомплектаций КАК Д "
           "ГДЕ ВЫРАЗИТЬ(Д.Комментарий КАК Строка(200)) = &М И НЕ Д.ПометкаУдаления")
q2.SetParameter("М", MARKER)
r2 = q2.Execute().Выгрузить()
assert r2.Количество() == 1, f"по маркеру найдено {r2.Количество()} документов"
ref = r2.Получить(0).Ссылка
obj2 = ref.ПолучитьОбъект()
assert obj2.ТабличнаяЧастьОстатков.Количество() == n
assert obj2.СписаниеПоНормам.Количество() == npn
assert obj2.СписаниеСверхНормы.Количество() == nsn
print(f"Persist: документ №{num}, ТЧ на месте после перечтения")
print("ANALIZ PASS")
