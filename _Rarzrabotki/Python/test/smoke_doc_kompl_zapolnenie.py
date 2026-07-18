# -*- coding: utf-8 -*-
"""Smoke 2: РасчетКомплектаций — создание/заполнение целевых документов из таблиц списания.
Ручная правка количества + ручная строка понаднормы ДОЛЖНЫ дойти до документа «Додаткова».
Целевые документы get-or-create по маркеру, запись БЕЗ проведения, ничего не удаляем."""
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

MARKER = "SMOKE_DOC_KOMPL_v1"

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')


def find_by_comment(doc_type, marker):
    q = buh.NewObject("Запрос")
    q.Text = (f"ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Ссылка ИЗ Документ.{doc_type} КАК Д "
              "ГДЕ ВЫРАЗИТЬ(Д.Комментарий КАК Строка(200)) = &М И НЕ Д.ПометкаУдаления")
    q.SetParameter("М", marker)
    r = q.Execute().Выгрузить()
    return r.Получить(0).Ссылка if r.Количество() > 0 else None


def get_or_create_target(doc_type, marker, org):
    ref = find_by_comment(doc_type, marker)
    if ref is not None:
        return ref
    d = getattr(buh.Документы, doc_type).СоздатьДокумент()
    d.Дата = datetime.datetime.now().replace(microsecond=0)
    d.Организация = org
    d.Комментарий = marker
    d.Записать()  # Запись, НЕ проведение
    return d.Ссылка


# --- наш документ-расчёт ---
ref = find_by_comment("РасчетКомплектаций", MARKER)
assert ref is not None, "сначала прогнать smoke_doc_kompl_analiz.py"
doc = ref.ПолучитьОбъект()
assert doc.СписаниеСверхНормы.Количество() > 0

# организация: из первой строки остатков её нет — возьмём первую по алфавиту непустую
qorg = buh.NewObject("Запрос")
qorg.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 О.Ссылка КАК Ссылка ИЗ Справочник.Организации КАК О ГДЕ НЕ О.ПометкаУдаления"
org = qorg.Execute().Выгрузить().Получить(0).Ссылка

# --- целевые документы (4 шт.) ---
kn_norm = get_or_create_target("КомплектацияНоменклатуры", MARKER + "_КН_НОРМ", org)
kn_over = get_or_create_target("КомплектацияНоменклатуры", MARKER + "_КН_ПОНАД", org)
mal_norm = get_or_create_target("ПередачаМалоценныхАктивовВЭксплуатацию", MARKER + "_МАЛ_НОРМ", org)
mal_over = get_or_create_target("ПередачаМалоценныхАктивовВЭксплуатацию", MARKER + "_МАЛ_ПОНАД", org)
print("Целевые:", kn_norm.Номер, kn_over.Номер, mal_norm.Номер, mal_over.Номер)

# --- прописать документы в строки складов ---
assert doc.ДокументиКомплектації.Количество() > 0, "нет строк ДокументиКомплектації (авто-склады)"
for i in range(doc.ДокументиКомплектації.Количество()):
    row = doc.ДокументиКомплектації.Получить(i)
    row.ДокументПоНормам = kn_norm
    row.ДокументДодаткова = kn_over
for i in range(doc.ДокументиМалоценки.Количество()):
    row = doc.ДокументиМалоценки.Получить(i)
    row.ДокументПоНормам = mal_norm
    row.ДокументДодаткова = mal_over

# --- ручная правка первой строки понаднормы ---
r0 = doc.СписаниеСверхНормы.Получить(0)
edited_nom = r0.Номенклатура
old_kol = float(r0.Количество)
new_kol = old_kol + 1
r0.Количество = new_kol
r0.Сумма = round(new_kol * float(r0.Цена), 2)
r0.Причина = "Тест: ручне коригування кількості"
edited_kol = new_kol
print(f"Правка: {edited_nom.Наименование} {old_kol} -> {new_kol}")

# --- ручная НОВАЯ строка понаднормы (автоподбор через экспортную функцию движка) ---
a0 = doc.ТабличнаяЧастьОстатков.Получить(0)
manual_nom, manual_skl = a0.Номенклатура, a0.Склад
dd = doc.ДанныеОстаткаДляСтроки(manual_nom, manual_skl)
assert dd is not None, "автоподбор не нашёл остаток"
nr = doc.СписаниеСверхНормы.Добавить()
nr.Номенклатура = manual_nom
nr.Склад = manual_skl
nr.Счет = dd.Счет
nr.Единица = dd.Единица
nr.ОбщееНазвание = dd.ОбщееНазвание
nr.Цена = dd.Цена
nr.Количество = 1
nr.Сумма = round(float(dd.Цена), 2)
nr.Причина = "Тест: ручний рядок"
print(f"Ручная строка: {manual_nom.Наименование}, счёт {dd.Счет.Код}, цена {dd.Цена}")

# --- заполнение (оба режима) ---
res = doc.ЗаполнитьДокументыПоАнализу(True, True)
print(f"Результат: Заповнено={res.Заповнено}, Пропущено={res.Пропущено}, "
      f"БезЕдиниці={res.БезЕдиниці}, БезНазначення={res.БезНазначення}, Помилок={res.Помилки.Количество()}")
for i in range(res.Помилки.Количество()):
    print("  Помилка:", res.Помилки.Получить(i))
assert res.Помилки.Количество() == 0
assert res.Заповнено >= 1

# --- проверки целевых ---
kn_over_obj = kn_over.ПолучитьОбъект()
n_over = kn_over_obj.Комплектующие.Количество()
found_edited = found_manual = False
scheta = set()
for i in range(n_over):
    r = kn_over_obj.Комплектующие.Получить(i)
    kod = r.СчетУчетаБУ.Код if not r.СчетУчетаБУ.Пустая() else ""
    assert kod != "", "пустой СчетУчетаБУ"
    scheta.add(str(kod))
    if str(r.Номенклатура.Наименование) == str(edited_nom.Наименование) and abs(float(r.Количество) - edited_kol) < 1e-6:
        found_edited = True
    if str(r.Номенклатура.Наименование) == str(manual_nom.Наименование) and abs(float(r.Количество) - 1) < 1e-6:
        found_manual = True
assert n_over > 0, "КН Додаткова пуста"
assert not any(k.startswith("22") for k in scheta), f"счета 22 в комплектации: {scheta}"
assert found_edited, "правка количества НЕ дошла до документа"
assert found_manual, "ручная строка НЕ дошла до документа"
print(f"КН Додаткова: {n_over} строк, счета {sorted(scheta)}, правка и ручная строка на месте")

kn_norm_obj = kn_norm.ПолучитьОбъект()
print(f"КН ПоНормам: {kn_norm_obj.Комплектующие.Количество()} строк")

mal_norm_obj = mal_norm.ПолучитьОбъект()
mal_over_obj = mal_over.ПолучитьОбъект()
n_mal = mal_norm_obj.МалоценныеАктивы.Количество() + mal_over_obj.МалоценныеАктивы.Количество()
if n_mal > 0:
    for obj in (mal_norm_obj, mal_over_obj):
        for i in range(obj.МалоценныеАктивы.Количество()):
            r = obj.МалоценныеАктивы.Получить(i)
            assert str(r.СчетУчетаБУ.Код).startswith("22"), f"не-22 счёт в малоценке: {r.СчетУчетаБУ.Код}"
            assert not r.НалоговоеНазначение.Пустая(), "пустое НалоговоеНазначение"
    print(f"Малоценка: {n_mal} строк, все счета 22*, налоговое назначение заполнено")
else:
    print("Малоценка: строк нет (на тестовом складе нет остатков сч.22) — маршрутизация не проверена данными")

# --- статус + повторное заполнение (идемпотентность) ---
st = buh.String(doc.Статус)
assert st in ("Документи створено", "Документы созданы"), st
res2 = doc.ЗаполнитьДокументыПоАнализу(True, True)
assert res2.Помилки.Количество() == 0
kn_over_obj2 = kn_over.ПолучитьОбъект()
assert kn_over_obj2.Комплектующие.Количество() == n_over, "повторное заполнение задвоило строки!"
print(f"Статус: {st}; повторное заполнение не задваивает ({n_over} строк)")

doc.Записать()
print("ZAPOLNENIE PASS")
