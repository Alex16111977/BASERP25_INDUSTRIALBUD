"""
Test: процедура СоздатьВедомостьНаВыплатуВКассуПоРаспределениюФ2 принимает
2 новых опциональных параметра (СтатьиДДСПоКлючу, СсылкаНаОтражениеЗП).

Сценарий:
1. Эталон Ф2 №000000026
2. Эталон ОтражениеЗПпоКазне №000000006
3. Построить словарь СтатьиДДСПоКлючу (Сотр_UID|ПодрОрг_UID → СтатьяДДС) из ТЧ.РаспределениеКазна
4. Вызвать Ф2Объект.СоздатьВедомостьНаВыплатуВКассуПоРаспределениюФ2(словарь, otr_ref)
5. Проверить ВКассу:
   - Шапка.А_ОтражениеЗПпоКазне = otr_ref ✓
   - Каждая строка А_Расшифровки имеет СтатьюДДС из словаря (по паре Сотр+ПодрОрг)
"""
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)


def find_doc(meta_name, number):
    q = erp.NewObject("Запрос")
    q.Текст = f'ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Ссылка ИЗ Документ.{meta_name} КАК Д ГДЕ Д.Номер = "{number}" И Д.Проведен'
    rs = q.Выполнить()
    if rs.Пустой(): return None
    sel = rs.Выбрать(); sel.Следующий()
    return sel.Ссылка


# 1. Эталоны
f2_ref = find_doc("РаспределениеФ2", "000000026")
otr_ref = find_doc("А_ОтражениеЗПпоКазне", "000000006")
if not f2_ref or not otr_ref:
    print("FAIL: эталоны не найдены")
    sys.exit(1)
f2_uid_str = erp.string(f2_ref.УникальныйИдентификатор())

# 2. Построить словарь СтатьиДДСПоКлючу из ОтражениеЗПпоКазне (только для строк нашего Ф2)
mod_priv = getattr(erp, "А_Привилегированный")
otr = otr_ref.ПолучитьОбъект()
slovar_py = {}
slovar_com = erp.NewObject("Соответствие")
for i in range(otr.РаспределениеКазна.Количество()):
    r = otr.РаспределениеКазна.Получить(i)
    if not r.ДокРаспределениеЗП or r.ДокРаспределениеЗП.Пустая(): continue
    if erp.string(r.ДокРаспределениеЗП.УникальныйИдентификатор()) != f2_uid_str: continue

    if r.Подразделение and not r.Подразделение.Пустая():
        podr_org = mod_priv.ОбеспечитьПодразделенияОрганизацийПоПодразделению(r.Подразделение)
    else:
        podr_org = None

    sotr_uid = erp.string(r.Сотрудник.УникальныйИдентификатор()) if r.Сотрудник and not r.Сотрудник.Пустая() else ""
    podr_uid = erp.string(podr_org.УникальныйИдентификатор()) if podr_org and not podr_org.Пустая() else ""
    klyuch = f"{sotr_uid}|{podr_uid}"
    slovar_com.Вставить(klyuch, r.СтатьяДвиженияДенежныхСредств)
    slovar_py[klyuch] = erp.string(r.СтатьяДвиженияДенежныхСредств.УникальныйИдентификатор()) if r.СтатьяДвиженияДенежныхСредств and not r.СтатьяДвиженияДенежныхСредств.Пустая() else ""

print(f"Словарь СтатьиДДСПоКлючу построен: {slovar_com.Количество()} пар (для Ф2 №000000026)")

# 3. УДАЛИТЬ существующую ВКассу для чистого старта (иначе она помнит предыдущие
#    значения А_ОтражениеЗПпоКазне — тест станет ложно-зелёным)
q = erp.NewObject("Запрос")
q.Текст = "ВЫБРАТЬ ПЕРВЫЕ 1 Вед.Ссылка КАК Ссылка ИЗ Документ.ВедомостьНаВыплатуЗарплатыВКассу КАК Вед ГДЕ Вед.А_РаспределениеФ2 = &Ф2"
q.УстановитьПараметр("Ф2", f2_ref)
rs = q.Выполнить()
if not rs.Пустой():
    sel = rs.Выбрать(); sel.Следующий()
    vk_old = sel.Ссылка.ПолучитьОбъект()
    if vk_old.Проведен:
        try:
            vk_old.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
        except Exception as e:
            info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
            print(f"  WARN: не удалось распровести: {info}")
    try:
        vk_old.ОбменДанными.Загрузка = True
        vk_old.Удалить()
        print("  Старая ВКассу удалена для чистого старта")
    except Exception as e:
        info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        print(f"  WARN: не удалось удалить: {info}")

    # Также очищаем обратную ссылку в Ф2
    f2_pre = f2_ref.ПолучитьОбъект()
    f2_pre.А_ВедомостьВКассу = erp.Документы.ВедомостьНаВыплатуЗарплатыВКассу.ПустаяСсылка()
    f2_pre.ОбменДанными.Загрузка = True
    f2_pre.Записать()

# 4. Вызвать процедуру с НОВЫМИ параметрами
f2_obj = f2_ref.ПолучитьОбъект()
try:
    f2_obj.СоздатьВедомостьНаВыплатуВКассуПоРаспределениюФ2(slovar_com, otr_ref)
    print("OK: процедура приняла 2 параметра без ошибки сигнатуры")
except Exception as e:
    info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
    print(f"FAIL: вызов с параметрами упал: {info}")
    sys.exit(1)

# 5. Перечитать Ф2 → получить ВКассу
f2_re = f2_ref.ПолучитьОбъект()
vk_ref = f2_re.А_ВедомостьВКассу
if not vk_ref or vk_ref.Пустая():
    print("FAIL: А_ВедомостьВКассу не заполнена")
    sys.exit(1)
vk = vk_ref.ПолучитьОбъект()

# 6. Проверка: ВКассу.А_ОтражениеЗПпоКазне = otr_ref
if not vk.А_ОтражениеЗПпоКазне or vk.А_ОтражениеЗПпоКазне.Пустая():
    print(f"FAIL: ВКассу.А_ОтражениеЗПпоКазне ПУСТО — параметр СсылкаНаОтражениеЗП не сработал")
    sys.exit(1)
vk_otr_uid = erp.string(vk.А_ОтражениеЗПпоКазне.УникальныйИдентификатор())
otr_uid = erp.string(otr_ref.УникальныйИдентификатор())
if vk_otr_uid != otr_uid:
    print(f"FAIL: ВКассу.А_ОтражениеЗПпоКазне UID не совпал: '{vk_otr_uid}' != '{otr_uid}'")
    sys.exit(1)
print(f"  Шапка.А_ОтражениеЗПпоКазне ✓")

# 7. Проверка: СтатьяДДС в А_Расшифровке совпадает со словарём
match = 0
mismatch = 0
no_expected = 0
for i in range(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Количество()):
    rr = vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Получить(i)
    sotr_uid = erp.string(rr.Сотрудник.УникальныйИдентификатор()) if rr.Сотрудник and not rr.Сотрудник.Пустая() else ""
    podr_uid = erp.string(rr.Подразделение.УникальныйИдентификатор()) if rr.Подразделение and not rr.Подразделение.Пустая() else ""
    klyuch = f"{sotr_uid}|{podr_uid}"
    expected_statya_uid = slovar_py.get(klyuch)

    actual_statya_uid = erp.string(rr.СтатьяДвиженияДенежныхСредств.УникальныйИдентификатор()) if rr.СтатьяДвиженияДенежныхСредств and not rr.СтатьяДвиженияДенежныхСредств.Пустая() else ""

    if expected_statya_uid is None:
        no_expected += 1
        continue
    if actual_statya_uid == expected_statya_uid:
        match += 1
    else:
        mismatch += 1
        print(f"  ❌ {rr.Сотрудник.Наименование}/{rr.Подразделение.Наименование}: ожидали UID='{expected_statya_uid}', получили '{actual_statya_uid}'")

print(f"\n  Расшифровка: match={match}, mismatch={mismatch}, нет в словаре={no_expected}")
if mismatch > 0:
    print("FAIL: есть несоответствия СтатьиДДС")
    sys.exit(1)
if match == 0:
    print("FAIL: ни одной матчевой строки — параметр СтатьиДДСПоКлючу не сработал")
    sys.exit(1)
print("\nPASS: параметры процедуры работают корректно")
