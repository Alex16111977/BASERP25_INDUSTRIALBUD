"""
Idempotency: дважды вызвать СоздатьВедомостиВКассуПоОтражениюЗП на эталоне.
Проверки:
- Кол. ВКассу с А_ОтражениеЗПпоКазне = эталон не меняется
- UUID каждой ВКассу не меняется между прогонами
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


def get_vks_uids():
    q2 = erp.NewObject("Запрос")
    q2.Текст = '''ВЫБРАТЬ Вед.Ссылка КАК Ссылка ИЗ Документ.ВедомостьНаВыплатуЗарплатыВКассу КАК Вед
    ГДЕ Вед.А_ОтражениеЗПпоКазне = &Отр И НЕ Вед.ПометкаУдаления
    УПОРЯДОЧИТЬ ПО Вед.Дата'''
    q2.УстановитьПараметр("Отр", otr_ref)
    rs2 = q2.Выполнить()
    uids = []
    sel2 = rs2.Выбрать()
    while sel2.Следующий():
        uids.append(erp.string(sel2.Ссылка.УникальныйИдентификатор()))
    return uids


# Первый вызов
otr_ref.ПолучитьОбъект().СоздатьВедомостиВКассуПоОтражениюЗП()
uids_1 = get_vks_uids()
print(f"Прогон 1: {len(uids_1)} ВКассу")

# Второй вызов
otr_ref.ПолучитьОбъект().СоздатьВедомостиВКассуПоОтражениюЗП()
uids_2 = get_vks_uids()
print(f"Прогон 2: {len(uids_2)} ВКассу")

if len(uids_1) != len(uids_2):
    print(f"FAIL: кол. ВКассу изменилось: {len(uids_1)} → {len(uids_2)}")
    sys.exit(1)
if sorted(uids_1) != sorted(uids_2):
    print(f"FAIL: UID-ы изменились между прогонами")
    print(f"  было: {sorted(uids_1)}")
    print(f"  стало: {sorted(uids_2)}")
    sys.exit(1)
print("\nPASS: идемпотентно — UID-ы и количество не меняются")
