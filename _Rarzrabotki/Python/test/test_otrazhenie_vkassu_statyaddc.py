"""
СтатьяДДС: для каждой строки А_Расшифровки во всех созданных ВКассу проверить
что СтатьяДвиженияДенежныхСредств точно совпадает с источником из
ТЧ.РаспределениеКазна по паре (Сотрудник, ПодрОрг).
"""
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)

q = erp.NewObject("Запрос")
q.Текст = '''ВЫБРАТЬ ПЕРВЫЕ 1 О.Ссылка КАК Ссылка ИЗ Документ.А_ОтражениеЗПпоКазне КАК О
ГДЕ О.Номер = "000000006" И О.Проведен'''
sel = q.Выполнить().Выбрать(); sel.Следующий()
otr_ref = sel.Ссылка

# Гарантируем актуальные ВКассу (Find-Or-Create обновит существующие)
otr_ref.ПолучитьОбъект().СоздатьВедомостиВКассуПоОтражениюЗП()

# Построить эталонный словарь Ф2 → {(Сотр_UID|ПодрОрг_UID): СтатьяДДС_UID}
mod_priv = getattr(erp, "А_Привилегированный")
otr = otr_ref.ПолучитьОбъект()
expected = {}
for i in range(otr.РаспределениеКазна.Количество()):
    r = otr.РаспределениеКазна.Получить(i)
    if not r.ДокРаспределениеЗП or r.ДокРаспределениеЗП.Пустая(): continue
    if str(r.ДокРаспределениеЗП.Метаданные().Имя) != "РаспределениеФ2": continue

    if r.Подразделение and not r.Подразделение.Пустая():
        podr_org = mod_priv.ОбеспечитьПодразделенияОрганизацийПоПодразделению(r.Подразделение)
    else:
        podr_org = None
    sotr_uid = erp.string(r.Сотрудник.УникальныйИдентификатор()) if r.Сотрудник and not r.Сотрудник.Пустая() else ""
    podr_uid = erp.string(podr_org.УникальныйИдентификатор()) if podr_org and not podr_org.Пустая() else ""
    key = f"{sotr_uid}|{podr_uid}"
    statya_uid = erp.string(r.СтатьяДвиженияДенежныхСредств.УникальныйИдентификатор()) if r.СтатьяДвиженияДенежныхСредств and not r.СтатьяДвиженияДенежныхСредств.Пустая() else ""
    f2_uid = erp.string(r.ДокРаспределениеЗП.УникальныйИдентификатор())
    expected.setdefault(f2_uid, {})[key] = statya_uid

print(f"Эталонный словарь: {sum(len(v) for v in expected.values())} пар по {len(expected)} Ф2")

# Перебрать все созданные ВКассу
q2 = erp.NewObject("Запрос")
q2.Текст = '''ВЫБРАТЬ Вед.Ссылка КАК Ссылка ИЗ Документ.ВедомостьНаВыплатуЗарплатыВКассу КАК Вед
ГДЕ Вед.А_ОтражениеЗПпоКазне = &Отр И НЕ Вед.ПометкаУдаления'''
q2.УстановитьПараметр("Отр", otr_ref)
rs2 = q2.Выполнить()
sel2 = rs2.Выбрать()

total_match = 0
total_mismatch = 0
total_no_expected = 0
total_vks = 0

while sel2.Следующий():
    total_vks += 1
    vk = sel2.Ссылка.ПолучитьОбъект()
    f2_uid = erp.string(vk.А_РаспределениеФ2.УникальныйИдентификатор())
    exp_map = expected.get(f2_uid, {})

    for i in range(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Количество()):
        rr = vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Получить(i)
        sotr_uid = erp.string(rr.Сотрудник.УникальныйИдентификатор()) if rr.Сотрудник and not rr.Сотрудник.Пустая() else ""
        podr_uid = erp.string(rr.Подразделение.УникальныйИдентификатор()) if rr.Подразделение and not rr.Подразделение.Пустая() else ""
        key = f"{sotr_uid}|{podr_uid}"

        actual_statya_uid = erp.string(rr.СтатьяДвиженияДенежныхСредств.УникальныйИдентификатор()) if rr.СтатьяДвиженияДенежныхСредств and not rr.СтатьяДвиженияДенежныхСредств.Пустая() else ""
        exp_statya_uid = exp_map.get(key)

        if exp_statya_uid is None:
            total_no_expected += 1
            print(f"  WARN [{vk.Номер}] {rr.Сотрудник.Наименование}/{rr.Подразделение.Наименование}: нет в ОтражениеЗПпоКазне")
            continue
        if actual_statya_uid == exp_statya_uid:
            total_match += 1
        else:
            total_mismatch += 1
            print(f"  ❌ [{vk.Номер}] {rr.Сотрудник.Наименование}/{rr.Подразделение.Наименование}: ожидали '{exp_statya_uid}', получили '{actual_statya_uid}'")

print(f"\nПросмотрено {total_vks} ВКассу. Итого строк А_Расшифровки:")
print(f"  match={total_match}, mismatch={total_mismatch}, нет в ОтражениеЗПпоКазне={total_no_expected}")
if total_mismatch > 0:
    print("FAIL: есть несоответствия СтатьиДДС")
    sys.exit(1)
if total_match == 0:
    print("FAIL: ни одной матчевой строки — что-то не так с ключами")
    sys.exit(1)
print("PASS: СтатьиДДС соответствуют источнику в ОтражениеЗПпоКазне")
