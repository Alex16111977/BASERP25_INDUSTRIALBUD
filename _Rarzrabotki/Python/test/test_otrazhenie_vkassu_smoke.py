"""
Smoke: вызвать новую процедуру СоздатьВедомостиВКассуПоОтражениюЗП на эталоне.
Проверки:
- Создано N ВКассу по числу уникальных Ф2 из ТЧ.РаспределениеКазна (= 7)
- Каждая ВКассу.А_ОтражениеЗПпоКазне = эталон ОтражениеЗПпоКазне
- Каждая ВКассу.А_РаспределениеФ2 = соответствующий Ф2
- Все ВКассу проведены
- Σ ТЧ Зарплата каждой ВКассу > 0
"""
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect(CONN)

# Эталон
q = erp.NewObject("Запрос")
q.Текст = '''ВЫБРАТЬ ПЕРВЫЕ 1 О.Ссылка КАК Ссылка
ИЗ Документ.А_ОтражениеЗПпоКазне КАК О
ГДЕ О.Номер = "000000006" И О.Проведен'''
rs = q.Выполнить()
sel = rs.Выбрать(); sel.Следующий()
otr_ref = sel.Ссылка

# Уникальные Ф2 из ТЧ.РаспределениеКазна
otr = otr_ref.ПолучитьОбъект()
unique_f2_uids = set()
for i in range(otr.РаспределениеКазна.Количество()):
    r = otr.РаспределениеКазна.Получить(i)
    if not r.ДокРаспределениеЗП or r.ДокРаспределениеЗП.Пустая(): continue
    if str(r.ДокРаспределениеЗП.Метаданные().Имя) != "РаспределениеФ2": continue
    unique_f2_uids.add(erp.string(r.ДокРаспределениеЗП.УникальныйИдентификатор()))

expected_count = len(unique_f2_uids)
print(f"Ожидаем создание {expected_count} ВКассу (по числу уникальных Ф2)")

# Вызвать новую процедуру
try:
    otr.СоздатьВедомостиВКассуПоОтражениюЗП()
    print("OK: процедура отработала без исключения")
except Exception as e:
    info = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
    print(f"FAIL: процедура упала: {info}")
    sys.exit(1)

# Проверить созданные ВКассу
q2 = erp.NewObject("Запрос")
q2.Текст = '''ВЫБРАТЬ Вед.Ссылка КАК Ссылка, Вед.Проведен КАК Проведен,
    Вед.А_РаспределениеФ2 КАК Ф2, Вед.А_ОтражениеЗПпоКазне КАК ОтрЗП,
    Вед.СуммаПоДокументу КАК S
ИЗ Документ.ВедомостьНаВыплатуЗарплатыВКассу КАК Вед
ГДЕ Вед.А_ОтражениеЗПпоКазне = &Отр И НЕ Вед.ПометкаУдаления'''
q2.УстановитьПараметр("Отр", otr_ref)
rs2 = q2.Выполнить()
sel2 = rs2.Выбрать()

found = []
while sel2.Следующий():
    found.append({
        "ref": sel2.Ссылка,
        "proveden": bool(sel2.Проведен),
        "f2_uid": erp.string(sel2.Ф2.УникальныйИдентификатор()) if sel2.Ф2 and not sel2.Ф2.Пустая() else "",
        "otr_uid": erp.string(sel2.ОтрЗП.УникальныйИдентификатор()) if sel2.ОтрЗП and not sel2.ОтрЗП.Пустая() else "",
        "sum": float(sel2.S or 0)
    })

print(f"\nНайдено ВКассу с А_ОтражениеЗПпоКазне = эталон: {len(found)}")
for fk in found:
    print(f"  {fk['ref']}: Провед={fk['proveden']}, Σ={fk['sum']:,.2f}")

# Проверки
otr_uid = erp.string(otr_ref.УникальныйИдентификатор())
fail = False

if len(found) != expected_count:
    print(f"FAIL: ожидали {expected_count} ВКассу, нашли {len(found)}")
    fail = True

for fk in found:
    if fk["otr_uid"] != otr_uid:
        print(f"FAIL: {fk['ref']} имеет неверный А_ОтражениеЗПпоКазне")
        fail = True
    if fk["f2_uid"] not in unique_f2_uids:
        print(f"FAIL: {fk['ref']} имеет А_РаспределениеФ2 не из эталона")
        fail = True
    if not fk["proveden"]:
        print(f"FAIL: {fk['ref']} НЕ проведена")
        fail = True
    if fk["sum"] <= 0:
        print(f"FAIL: {fk['ref']} имеет нулевую сумму")
        fail = True

if fail:
    sys.exit(1)

print(f"\nPASS: smoke test пройден ({len(found)} ВКассу, все проведены)")
