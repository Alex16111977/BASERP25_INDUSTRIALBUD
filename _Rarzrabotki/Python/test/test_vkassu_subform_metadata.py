"""
Smoke метаданных: проверяет что:
1. Форма Документ.ВедомостьНаВыплатуЗарплатыВКассу.Форма.А_ФормаСпискаРасшифровкиПоФЛ
   зарегистрирована в составе документа.
2. ТЧ А_РасшифровкаВыплатыЗарплатаПоФизлицам существует с нужными реквизитами.
3. Реквизит А_ОтражениеЗПпоКазне существует на документе (от предыдущих коммитов).
"""
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)

# 1. Документ существует
meta_doc = erp.Метаданные.Документы.Найти("ВедомостьНаВыплатуЗарплатыВКассу")
if meta_doc is None:
    print("FAIL: документ ВедомостьНаВыплатуЗарплатыВКассу не найден")
    sys.exit(1)
print(f"OK: документ {meta_doc.Имя} найден")

# 2. Форма зарегистрирована
form_meta = None
form_names = []
for i in range(meta_doc.Формы.Количество()):
    f = meta_doc.Формы.Получить(i)
    form_names.append(str(f.Имя))
    if str(f.Имя) == "А_ФормаСпискаРасшифровкиПоФЛ":
        form_meta = f

if form_meta is None:
    print(f"FAIL: форма А_ФормаСпискаРасшифровкиПоФЛ не найдена. Известные формы: {form_names}")
    sys.exit(1)
print(f"OK: форма '{form_meta.Имя}' зарегистрирована (всего форм у документа: {len(form_names)})")

# 3. ТЧ А_РасшифровкаВыплатыЗарплатаПоФизлицам существует с нужными реквизитами
tch = None
for i in range(meta_doc.ТабличныеЧасти.Количество()):
    t = meta_doc.ТабличныеЧасти.Получить(i)
    if str(t.Имя) == "А_РасшифровкаВыплатыЗарплатаПоФизлицам":
        tch = t
        break

if tch is None:
    print("FAIL: ТЧ А_РасшифровкаВыплатыЗарплатаПоФизлицам не найдена")
    sys.exit(1)
print(f"OK: ТЧ {tch.Имя} существует")

required_attrs = ["ИдентификаторСтроки", "Сотрудник", "ФизическоеЛицо", "Подразделение",
                  "ПериодВзаиморасчетов", "СтатьяФинансирования", "СтатьяРасходов",
                  "ДокументОснование", "КВыплате", "КомпенсацияЗаЗадержкуЗарплаты",
                  "ГруппаУчетаНачислений", "НаправлениеДеятельности",
                  "СтатьяДвиженияДенежныхСредств"]
tch_attr_names = []
for i in range(tch.Реквизиты.Количество()):
    tch_attr_names.append(str(tch.Реквизиты.Получить(i).Имя))

missing = [a for a in required_attrs if a not in tch_attr_names]
if missing:
    print(f"FAIL: ТЧ не имеет реквизитов: {missing}")
    sys.exit(1)
print(f"OK: ТЧ имеет все {len(required_attrs)} нужных реквизитов")

# 4. Реквизит А_ОтражениеЗПпоКазне (на шапке)
hdr_attr_names = []
for i in range(meta_doc.Реквизиты.Количество()):
    hdr_attr_names.append(str(meta_doc.Реквизиты.Получить(i).Имя))

if "А_ОтражениеЗПпоКазне" not in hdr_attr_names:
    print("FAIL: реквизит шапки А_ОтражениеЗПпоКазне не найден")
    sys.exit(1)
print("OK: реквизит шапки А_ОтражениеЗПпоКазне существует")

if "А_РаспределениеФ2" not in hdr_attr_names:
    print("FAIL: реквизит шапки А_РаспределениеФ2 не найден")
    sys.exit(1)
print("OK: реквизит шапки А_РаспределениеФ2 существует")

print("\nPASS: метаданные документа и формы корректны")
