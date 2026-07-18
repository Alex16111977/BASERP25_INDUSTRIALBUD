# -*- coding: utf-8 -*-
# Діагностика: як знайти документ ERP за даними з Казни?
# Гіпотези:
#   1. UUID зберігається при обміні 1:1 (прямий ПолучитьСсылку)
#   2. Регістр СоответствияОбъектовИнформационныхБаз (БСП) містить маппінг
#   3. Пошук за Номер + Дата
#
# ВАЖЛИВО: параметри дат передаємо через pywintypes.Time (datetime → COM-дата)

import win32com.client, sys, pywintypes, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
kazna = v8.Connect('Srvr="localhost";Ref="kazna";Usr="cfo";Pwd="2442"')
erp   = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')


def com_date(y, m, d, hh=0, mm=0, ss=0):
    return pywintypes.Time(datetime.datetime(y, m, d, hh, mm, ss))


print("=" * 80)
print("КРОК 1: беремо 5 документів РаспределениеЗаработнойПлаты з Казни за грудень 2025")
print("=" * 80)

q = kazna.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ ПЕРВЫЕ 5
    Д.Ссылка КАК Ссылка,
    Д.Номер КАК Номер,
    Д.Дата КАК Дата,
    Д.КассаДатаС КАК КассаДатаС,
    Д.КассаДатаПо КАК КассаДатаПо,
    Д.Организация.Наименование КАК Орг,
    Д.Организация.А_ПереносЕРП КАК ПереносЕРП
ИЗ Документ.РаспределениеЗаработнойПлаты КАК Д
ГДЕ Д.КассаДатаС >= &Н И Д.КассаДатаПо <= &К
    И Д.Организация.А_ПереносЕРП
УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ
"""
q.SetParameter("Н", com_date(2025, 12, 1))
q.SetParameter("К", com_date(2025, 12, 31, 23, 59, 59))
sel = q.Execute().Choose()

docs_rasp = []
while sel.Next():
    uuid_kazna = kazna.XMLСтрока(sel.Ссылка.УникальныйИдентификатор())
    docs_rasp.append({
        "type": "РаспределениеЗаработнойПлаты",
        "erp_type": "А_РаспределениеЗаработнойПлаты",
        "uuid": uuid_kazna,
        "num": kazna.String(sel.Номер),
        "date": kazna.String(sel.Дата),
        "орг": kazna.String(sel.Орг),
    })

print(f"\nЗнайдено {len(docs_rasp)} РаспределениеЗаработнойПлаты:")
for d in docs_rasp:
    print(f"  Казна UUID={d['uuid']}  Номер={d['num']}  Дата={d['date']}  Орг={d['орг']}")

print()
print("=" * 80)
print("КРОК 2: беремо 5 документів РаспределениеФ2 з Казни за грудень 2025")
print("=" * 80)

q2 = kazna.NewObject("Запрос")
q2.Text = """
ВЫБРАТЬ ПЕРВЫЕ 5
    Д.Ссылка КАК Ссылка,
    Д.Номер КАК Номер,
    Д.Дата КАК Дата,
    Д.КассаДатаС КАК КассаДатаС,
    Д.КассаДатаПо КАК КассаДатаПо,
    Д.Организация.Наименование КАК Орг
ИЗ Документ.РаспределениеФ2 КАК Д
ГДЕ Д.КассаДатаС >= &Н И Д.КассаДатаПо <= &К
УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ
"""
q2.SetParameter("Н", com_date(2025, 12, 1))
q2.SetParameter("К", com_date(2025, 12, 31, 23, 59, 59))
sel2 = q2.Execute().Choose()

docs_f2 = []
while sel2.Next():
    uuid_kazna = kazna.XMLСтрока(sel2.Ссылка.УникальныйИдентификатор())
    docs_f2.append({
        "type": "РаспределениеФ2",
        "erp_type": "РаспределениеФ2",
        "uuid": uuid_kazna,
        "num": kazna.String(sel2.Номер),
        "date": kazna.String(sel2.Дата),
        "орг": kazna.String(sel2.Орг),
    })

print(f"\nЗнайдено {len(docs_f2)} РаспределениеФ2:")
for d in docs_f2:
    print(f"  Казна UUID={d['uuid']}  Номер={d['num']}  Дата={d['date']}  Орг={d['орг']}")

all_docs = docs_rasp + docs_f2

print()
print("=" * 80)
print("КРОК 3: перевіряємо кожен документ в ERP — 3 гіпотези")
print("=" * 80)

for d in all_docs:
    print(f"\n--- {d['type']} №{d['num']} від {d['date']} ---")
    print(f"    Казна UUID: {d['uuid']}")

    # Гіпотеза 1: прямий пошук за UUID (Конвертация зберегла UUID 1:1)
    h1 = "?"
    try:
        erp_doc_mgr = getattr(erp.Документы, d['erp_type'])
        uid = erp.NewObject("УникальныйИдентификатор", d['uuid'])
        ref_by_uuid = erp_doc_mgr.ПолучитьСсылку(uid)
        obj = ref_by_uuid.ПолучитьОбъект()
        if obj is not None:
            h1 = f"OK  {erp.String(ref_by_uuid)}"
        else:
            h1 = "FAIL (бите посилання)"
    except Exception as e:
        h1 = f"EXC: {e}"
    print(f"    [Г1] Прямий UUID:                  {h1}")

    # Гіпотеза 2: регістр СоответствияОбъектовИнформационныхБаз
    # Спроба A: UUID з Казни як рядок у УникальныйИдентификаторПриемника
    h2a = "?"
    try:
        qe = erp.NewObject("Запрос")
        qe.Text = """
        ВЫБРАТЬ
            СО.УникальныйИдентификаторИсточника КАК СсылкаERP,
            СО.ТипПриемника КАК ТипПриемника,
            СО.ТипИсточника КАК ТипИсточника,
            СО.УзелИнформационнойБазы КАК Узел
        ИЗ РегистрСведений.СоответствияОбъектовИнформационныхБаз КАК СО
        ГДЕ СО.УникальныйИдентификаторПриемника = &UUID
        """
        qe.SetParameter("UUID", d['uuid'])
        r = qe.Execute().Выгрузить()
        if r.Количество() == 0:
            h2a = "FAIL (нема запису)"
        else:
            rows = []
            for i in range(r.Количество()):
                row = r.Получить(i)
                rows.append(f"{erp.String(row.СсылкаERP)} [ТипПр={erp.String(row.ТипПриемника)}, ТипИст={erp.String(row.ТипИсточника)}, Узел={erp.String(row.Узел)}]")
            h2a = "OK " + " | ".join(rows)
    except Exception as e:
        h2a = f"EXC: {e}"
    print(f"    [Г2a] СОИБ: UUID у Приемник:       {h2a}")

    # Спроба B: UUID з Казни як рядок у ресурсі УникальныйИдентификаторИсточникаСтрокой
    h2b = "?"
    try:
        qe = erp.NewObject("Запрос")
        qe.Text = """
        ВЫБРАТЬ
            СО.УникальныйИдентификаторИсточника КАК СсылкаERP,
            СО.УникальныйИдентификаторПриемника КАК UUIDПриемника,
            СО.ТипПриемника КАК ТипПриемника,
            СО.ТипИсточника КАК ТипИсточника,
            СО.УзелИнформационнойБазы КАК Узел
        ИЗ РегистрСведений.СоответствияОбъектовИнформационныхБаз КАК СО
        ГДЕ СО.УникальныйИдентификаторИсточникаСтрокой = &UUID
        """
        qe.SetParameter("UUID", d['uuid'])
        r = qe.Execute().Выгрузить()
        if r.Количество() == 0:
            h2b = "FAIL (нема запису)"
        else:
            rows = []
            for i in range(r.Количество()):
                row = r.Получить(i)
                rows.append(f"{erp.String(row.СсылкаERP)} [ERPUUID={erp.String(row.UUIDПриемника)[:20]}..., Узел={erp.String(row.Узел)}]")
            h2b = "OK " + " | ".join(rows)
    except Exception as e:
        h2b = f"EXC: {e}"
    print(f"    [Г2b] СОИБ: UUID у ИсточникаСтрокой: {h2b}")

    # Гіпотеза 3: пошук за Номер + Дата (день)
    h3 = "?"
    try:
        qe = erp.NewObject("Запрос")
        qe.Text = f"""
        ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка, Д.Дата
        ИЗ Документ.{d['erp_type']} КАК Д
        ГДЕ Д.Номер = &Н И Д.Дата МЕЖДУ &НД И &КД
        """
        qe.SetParameter("Н", d['num'])
        date_parts = d['date'].split(" ")[0].split(".")
        y, m, dd = int(date_parts[2]), int(date_parts[1]), int(date_parts[0])
        qe.SetParameter("НД", com_date(y, m, dd))
        qe.SetParameter("КД", com_date(y, m, dd, 23, 59, 59))
        r = qe.Execute().Выгрузить()
        if r.Количество() == 0:
            h3 = "FAIL"
        else:
            h3 = f"OK  {erp.String(r.Получить(0).Ссылка)}"
    except Exception as e:
        h3 = f"EXC: {e}"
    print(f"    [Г3] Номер+Дата:                   {h3}")

print()
print("=" * 80)
print("КРОК 4: загальна статистика по регістру СоответствияОбъектовИнформационныхБаз")
print("=" * 80)

try:
    qe = erp.NewObject("Запрос")
    qe.Text = """
    ВЫБРАТЬ
        ВЫРАЗИТЬ(СО.УникальныйИдентификаторИсточника КАК Документ.А_РаспределениеЗаработнойПлаты).Ссылка КАК СсылкаРасп,
        СО.ТипПриемника, СО.ТипИсточника, СО.УзелИнформационнойБазы, СО.УникальныйИдентификаторПриемника
    ИЗ РегистрСведений.СоответствияОбъектовИнформационныхБаз КАК СО
    ГДЕ ТИПЗНАЧЕНИЯ(СО.УникальныйИдентификаторИсточника) = ТИП(Документ.А_РаспределениеЗаработнойПлаты)
    """
    r = qe.Execute().Выгрузить()
    print(f"\n  Записів для А_РаспределениеЗаработнойПлаты: {r.Количество()}")
    for i in range(min(3, r.Количество())):
        row = r.Получить(i)
        print(f"    {erp.String(row.СсылкаРасп)[:50]}")
        print(f"      ТипПр={erp.String(row.ТипПриемника)} | ТипИст={erp.String(row.ТипИсточника)}")
        print(f"      UUIDПрием={erp.String(row.УникальныйИдентификаторПриемника)}")
        print(f"      Узел={erp.String(row.УзелИнформационнойБазы)}")
except Exception as e:
    print(f"  EXC: {e}")

try:
    qe = erp.NewObject("Запрос")
    qe.Text = """
    ВЫБРАТЬ
        ВЫРАЗИТЬ(СО.УникальныйИдентификаторИсточника КАК Документ.РаспределениеФ2).Ссылка КАК СсылкаФ2,
        СО.ТипПриемника, СО.ТипИсточника, СО.УзелИнформационнойБазы, СО.УникальныйИдентификаторПриемника
    ИЗ РегистрСведений.СоответствияОбъектовИнформационныхБаз КАК СО
    ГДЕ ТИПЗНАЧЕНИЯ(СО.УникальныйИдентификаторИсточника) = ТИП(Документ.РаспределениеФ2)
    """
    r = qe.Execute().Выгрузить()
    print(f"\n  Записів для РаспределениеФ2: {r.Количество()}")
    for i in range(min(3, r.Количество())):
        row = r.Получить(i)
        print(f"    {erp.String(row.СсылкаФ2)[:50]}")
        print(f"      ТипПр={erp.String(row.ТипПриемника)} | ТипИст={erp.String(row.ТипИсточника)}")
        print(f"      UUIDПрием={erp.String(row.УникальныйИдентификаторПриемника)}")
        print(f"      Узел={erp.String(row.УзелИнформационнойБазы)}")
except Exception as e:
    print(f"  EXC: {e}")

print()
print("=" * 80)
print("ВИСНОВОК: обирай гіпотезу, де всі перевірки = OK")
print("=" * 80)
