# -*- coding: utf-8 -*-
"""Smoke 5: РасчетКомплектаций — флаг РасхождениеЕдиниц + печать «Розбіжності одиниць».
Наивная перекрёстная проверка правила (Python) против флага движка (BSL); отчёт содержит
только эталоны с расхождениями. Get-or-create, без удаления."""
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MARKER = "SMOKE_DOC_KOMPL_v1"

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

q = buh.NewObject("Запрос")
q.Text = ("ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Ссылка ИЗ Документ.РасчетКомплектаций КАК Д "
          "ГДЕ ВЫРАЗИТЬ(Д.Комментарий КАК Строка(200)) = &М И НЕ Д.ПометкаУдаления")
q.SetParameter("М", MARKER)
r = q.Execute().Выгрузить()
if r.Количество() > 0:
    doc = r.Получить(0).Ссылка.ПолучитьОбъект()
else:
    # создать заново (прежний тестовый мог быть помечен на удаление пользователем)
    doc = buh.Документы.РасчетКомплектаций.СоздатьДокумент()
    doc.Дата = datetime.datetime.now().replace(microsecond=0)
    doc.Заполнить(None)
    doc.Комментарий = MARKER
    doc.Спецификация = buh.Справочники.СтруктураСебестоимости.НайтиПоКоду("000000004")
    doc.Период = datetime.datetime(2026, 7, 7)
    sch = buh.NewObject("Массив")
    for kod in ("20", "22", "28"):
        sch.Добавить(buh.ПланыСчетов.Хозрасчетный.НайтиПоКоду(kod))
    qs = buh.NewObject("Запрос")
    qs.Text = """
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
    qs.SetParameter("НаДату", datetime.datetime(2026, 7, 7, 23, 59, 59))
    qs.SetParameter("Счета", sch)
    rows = qs.Execute().Выгрузить()
    chosen = [rows.Получить(i).Склад for i in range(rows.Количество())
              if "МХП" in rows.Получить(i).Склад.Наименование.upper()][:2] or [rows.Получить(0).Склад]
    doc.СкладыОстатков.Очистить()
    for s in chosen:
        doc.СкладыОстатков.Добавить().Склад = s

doc.РассчитатьАнализ()
n = doc.ТабличнаяЧастьОстатков.Количество()
assert n > 0


def uid(ref):
    return buh.String(ref.УникальныйИдентификатор())


# --- кэш единиц эталонов (наивный, отдельным запросом) ---
etalon_refs = {}
for i in range(n):
    row = doc.ТабличнаяЧастьОстатков.Получить(i)
    e = row.ОбщееНазвание
    if not e.Пустая():
        etalon_refs.setdefault(uid(e), e)
qe = buh.NewObject("Запрос")
qe.Text = ("ВЫБРАТЬ ОН.Ссылка КАК Эталон, ОН.Единица КАК Единица "
           "ИЗ Справочник.ОбщиеНазванияНоменклатуры КАК ОН ГДЕ ОН.Ссылка В (&Эталоны)")
arr = buh.NewObject("Массив")
for e in etalon_refs.values():
    arr.Добавить(e)
qe.SetParameter("Эталоны", arr)
unit_of_etalon = {}
vb = qe.Execute().Выбрать()
while vb.Следующий():
    # нормализация как в движке: СокрЛП(Наименование) — дубли " шт"/"шт" считаются одной единицей
    unit_of_etalon[uid(vb.Эталон)] = None if vb.Единица.Пустая() else str(vb.Единица.Наименование).strip()

# --- наивная проверка правила по каждой строке ---
mismatches = 0
flagged = 0
etalon_flag = {}   # uid эталона -> есть ли расхождение
for i in range(n):
    row = doc.ТабличнаяЧастьОстатков.Получить(i)
    ed_ss = None if row.ЕдиницаСС.Пустая() else str(row.ЕдиницаСС.Наименование).strip()
    ed_ost = None if row.Единица.Пустая() else str(row.Единица.Наименование).strip()
    has_etalon = not row.ОбщееНазвание.Пустая()
    ed_et = unit_of_etalon.get(uid(row.ОбщееНазвание)) if has_etalon else None
    if has_etalon and ed_et is None:
        expected = True
    else:
        filled = [x for x in (ed_ss, ed_ost, ed_et) if x is not None]
        expected = len(set(filled)) > 1
    actual = bool(row.РасхождениеЕдиниц)
    if actual != expected:
        mismatches += 1
        if mismatches <= 3:
            print(f"  MISMATCH стр{i}: {row.Номенклатура.Наименование}: движок={actual}, ожидание={expected}")
    if actual:
        flagged += 1
        if has_etalon:
            etalon_flag[uid(row.ОбщееНазвание)] = True
assert mismatches == 0, f"правило разошлось на {mismatches} строках"
etalons_with = len(etalon_flag)
print(f"Флаг: {flagged}/{n} строк с расхождением, эталонов с расхождением: {etalons_with}; правило совпало на всех строках")

# --- печать ---
td = doc.СформироватьПечатьРозбіжностіОдиниць()
h, w = td.ВысотаТаблицы, td.ШиринаТаблицы
assert h > 8 and w >= 5, f"печать {h}x{w}"
texts = []
for row in range(1, h + 1):
    for col in range(1, 6):
        t = td.Область(row, col, row, col).Текст
        if t:
            texts.append(t)
alltext = "\n".join(texts)
assert "Розбіжності одиниць" in alltext
itog_line = [t for t in texts if t.startswith("Еталонів з розбіжностями:")]
assert itog_line, "нет итоговой строки"
import re
m = re.search(r"Еталонів з розбіжностями: (\d+), номенклатур з розбіжністю: (\d+)", itog_line[0])
assert m, itog_line[0]
n_et, n_nom = int(m.group(1)), int(m.group(2))
assert n_et == etalons_with, f"в отчёте эталонов {n_et}, ожидалось {etalons_with}"
assert n_nom >= 1 or flagged == 0
print(f"Печать: {h}x{w}; итог '{itog_line[0]}' совпал с расчётом")

# эталон БЕЗ расхождений не должен попасть в отчёт
clean_etalon = None
for i in range(n):
    row = doc.ТабличнаяЧастьОстатков.Получить(i)
    if not row.ОбщееНазвание.Пустая() and uid(row.ОбщееНазвание) not in etalon_flag:
        clean_etalon = row.ОбщееНазвание.Наименование
        break
if clean_etalon is not None:
    assert not any(t.startswith(str(clean_etalon) + " [") for t in texts), f"чистый эталон в отчёте: {clean_etalon}"
    print(f"Чистый эталон «{clean_etalon}» в отчёт не попал — фильтрация верна")
else:
    print("Все эталоны с расхождениями (чистых нет) — фильтрацию подтвердить нечем")

doc.Записать()
print("EDINICY PASS")
