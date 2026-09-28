# -*- coding: utf-8 -*-
"""Приймання: контроль «той ли склад» у документах списання РасчетКомплектаций (BuhBud, 28.09.2026).

Промт: _Rarzrabotki/notebook/knowledge_Бухгалтерия_РасчетКомплектаций/PROMPT_dorabotka_KontrolSkladaDokumentov.md
Бойові РК НЕ записуються успішно, цільові документи НЕ перепроводяться: усі зміни — лише в памʼяті обʼєкта.
"""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
b = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

FAILS = []


def check(cond, msg):
    print(("PASS " if cond else "FAIL ") + msg)
    if not cond:
        FAILS.append(msg)


def query(text, **params):
    q = b.NewObject("Запрос")
    q.Text = text
    for k, v in params.items():
        q.SetParameter(k, v)
    return q.Execute().Выгрузить()


def uid(ref):
    # COM-посилання в Python порівнюються за ідентичністю обʼєкта — порівнюємо GUID
    return b.String(ref.УникальныйИдентификатор())


def rk_object(number):
    r = query('ВЫБРАТЬ Р.Ссылка КАК Ссылка ИЗ Документ.РасчетКомплектаций КАК Р '
              'ГДЕ Р.Номер = &Н И НЕ Р.ПометкаУдаления', Н=number)
    assert r.Количество() == 1, number
    return r.Получить(0).Ссылка.ПолучитьОбъект()


def mismatches(obj):
    arr = obj.РасхожденияСкладовДокументов()
    return [arr.Получить(i) for i in range(arr.Количество())]


def doc_snapshot(doc_number):
    r = query('ВЫБРАТЬ Д.Ссылка КАК Ссылка, Д.ВерсияДанных КАК Версия ИЗ Документ.ПередачаМалоценныхАктивовВЭксплуатацию КАК Д '
              'ГДЕ Д.Номер = &Н', Н=doc_number)
    s = r.Получить(0)
    rows = query('ВЫБРАТЬ Т.НомерСтроки КАК Н, Т.Номенклатура КАК Ном, Т.Количество КАК К '
                 'ИЗ Документ.ПередачаМалоценныхАктивовВЭксплуатацию.МалоценныеАктивы КАК Т '
                 'ГДЕ Т.Ссылка = &С УПОРЯДОЧИТЬ ПО Н', С=s.Ссылка)
    data = [(rows.Получить(i).Н, b.String(rows.Получить(i).Ном), float(rows.Получить(i).К))
            for i in range(rows.Количество())]
    return b.String(s.Версия), data


# 1. Відомі розбіжності: РК №10 і №11 — по 2 (ДокументиМалоценки, рядок 1, обидві колонки)
for num, sklad_doc in (("00000000010", "ІС МД ООН Админ №1 Виробничий"),
                       ("00000000011", "МД IRС 2026 30 м №3 Виробничий")):
    obj = rk_object(num)
    m = mismatches(obj)
    check(len(m) == 2, f"РК {num}: розбіжностей 2 (є {len(m)})")
    check(all(x.ИмяТЧ == "ДокументиМалоценки" and int(x.НомерСтроки) == 1 for x in m),
          f"РК {num}: розбіжності в ДокументиМалоценки, рядок 1")
    check(sorted(x.ИмяКолонки for x in m) == ["ДокументДодаткова", "ДокументПоНормам"],
          f"РК {num}: обидві колонки")
    check(all(sklad_doc in x.Текст for x in m), f"РК {num}: у тексті склад документа «{sklad_doc}»")

# 2. Інваріант по всіх РК: порожній результат ⇔ у кожного документа склад = склад рядка
all_rk = query('ВЫБРАТЬ Р.Ссылка КАК Ссылка ИЗ Документ.РасчетКомплектаций КАК Р ГДЕ НЕ Р.ПометкаУдаления')
total_bad = 0
for i in range(all_rk.Количество()):
    ref = all_rk.Получить(i).Ссылка
    obj = ref.ПолучитьОбъект()
    got = len(mismatches(obj))
    expected = 0
    for tch in ("ДокументиКомплектації", "ДокументиМалоценки"):
        t = getattr(obj, tch)
        for j in range(t.Количество()):
            row = t.Получить(j)
            for col in ("ДокументПоНормам", "ДокументДодаткова"):
                d = getattr(row, col)
                if not b.ЗначениеЗаполнено(d):
                    continue
                bad = uid(d.Склад) != uid(row.Склад) or (
                    b.ЗначениеЗаполнено(row.Организация) and uid(d.Организация) != uid(row.Организация))
                expected += 1 if bad else 0
    total_bad += got
    if got != expected:
        check(False, f"{b.String(ref)}: функція {got}, прямий перерахунок {expected}")
check(True, f"інваріант по {all_rk.Количество()} РК звірено, розбіжностей усього {total_bad}")
check(total_bad == 4, f"по базі рівно 4 розбіжності (РК №10 і №11), є {total_bad}")

# 3. Запис РК №10 відхиляється
obj = rk_object("00000000010")
b.ПолучитьСообщенияПользователю(True)
wrote = True
try:
    obj.Записать()
except Exception:
    wrote = False
msgs = b.ПолучитьСообщенияПользователю(True)
texts = [msgs.Получить(i).Текст for i in range(msgs.Количество())]
check(not wrote, "РК №10: Записать() відхилено")
check(any("ІБ00-000525" in t for t in texts), "РК №10: повідомлення містить ІБ00-000525")

# 4. «Заповнити документи» на РК №10 нічого не пише
before = doc_snapshot("ІБ00-000525")
obj = rk_object("00000000010")
res = obj.ЗаполнитьДокументыПоАнализу(True, True)
after = doc_snapshot("ІБ00-000525")
check(res.Помилки.Количество() == 2, f"Заповнення: 2 помилки (є {res.Помилки.Количество()})")
check(res.Заповнено == 0, "Заповнення: Заповнено = 0")
check(before == after, "ІБ00-000525: версія даних і рядки не змінились")

# 5. Очистили документи рядка 1 у памʼяті → розбіжностей немає
obj = rk_object("00000000010")
row = obj.ДокументиМалоценки.Получить(0)
row.ДокументПоНормам = b.Документы.ПередачаМалоценныхАктивовВЭксплуатацию.ПустаяСсылка()
row.ДокументДодаткова = b.Документы.ПередачаМалоценныхАктивовВЭксплуатацию.ПустаяСсылка()
check(len(mismatches(obj)) == 0, "РК №10 після очищення рядка 1 (у памʼяті): розбіжностей 0")

# 6. Проведення: лишаємо в памʼяті ЛИШЕ чужі документи → жоден не проводиться
obj = rk_object("00000000010")
obj.ДокументиКомплектації.Очистить()
for j in range(1, obj.ДокументиМалоценки.Количество()):
    r = obj.ДокументиМалоценки.Получить(j)
    r.ДокументПоНормам = b.Документы.ПередачаМалоценныхАктивовВЭксплуатацию.ПустаяСсылка()
    r.ДокументДодаткова = b.Документы.ПередачаМалоценныхАктивовВЭксплуатацию.ПустаяСсылка()
before = doc_snapshot("ІБ00-000525")
res = obj.ПровестиДокументы()
after = doc_snapshot("ІБ00-000525")
check(res.Оброблено == 0, f"Проведення: Оброблено = 0 (є {res.Оброблено})")
check(res.Пропущено == 2, f"Проведення: Пропущено = 2 (є {res.Пропущено})")
check(res.Помилки.Количество() == 2, f"Проведення: 2 помилки (є {res.Помилки.Количество()})")
check(before == after, "Проведення: ІБ00-000525 не перепроведено")

print()
print("ИТОГ:", "FAIL " + str(len(FAILS)) if FAILS else "ВСЕ PASS")
sys.exit(1 if FAILS else 0)
