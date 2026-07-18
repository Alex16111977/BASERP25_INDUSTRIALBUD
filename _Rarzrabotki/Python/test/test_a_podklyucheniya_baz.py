# -*- coding: utf-8 -*-
"""Приёмка А_ПодключенияБазСервер: сид, идемпотентность, соединения, негатив."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

def err_text(e):
    return e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
M = erp.А_ПодключенияБазСервер
fails = 0

# 1. Сид
M.ЗаполнитьНастройкиПоУмолчанию()
q = erp.NewObject("Запрос")
q.Text = "ВЫБРАТЬ ИдентификаторБазы, Сервер, ИмяБазы, Пользователь ИЗ РегистрСведений.А_НастройкиПодключенияБаз УПОРЯДОЧИТЬ ПО ИдентификаторБазы"
t = q.Execute().Выгрузить()
print(f"записей после сида: {t.Количество()} (ожидаем 4)")
if t.Количество() != 4:
    fails += 1
for i in range(t.Количество()):
    r = t.Получить(i)
    print(f"  {r.ИдентификаторБазы}: Srvr={r.Сервер} Ref={r.ИмяБазы} Usr={r.Пользователь}")

# 2. Идемпотентность: правим запись, сид повторно, значение не должно перетереться
rec = erp.РегистрыСведений.А_НастройкиПодключенияБаз.СоздатьМенеджерЗаписи()
rec.ИдентификаторБазы = "zup"
rec.Прочитать()
rec.Комментарий = "ТЕСТ-НЕ-ПЕРЕТИРАТЬ"
rec.Записать(True)
M.ЗаполнитьНастройкиПоУмолчанию()
rec2 = erp.РегистрыСведений.А_НастройкиПодключенияБаз.СоздатьМенеджерЗаписи()
rec2.ИдентификаторБазы = "zup"
rec2.Прочитать()
if rec2.Комментарий == "ТЕСТ-НЕ-ПЕРЕТИРАТЬ":
    print("OK идемпотентность сида")
else:
    print(f"FAIL идемпотентность: комментарий = {rec2.Комментарий}")
    fails += 1
rec2.Комментарий = "ЗУП регламентная (бухгалтерская зарплата)"
rec2.Записать(True)

# 3. Соединения со всеми 4 базами через новую функцию
for base in ["zup", "zup2", "bas_industrialbud", "kazna"]:
    try:
        conn = M.УстановитьСоединение(base)
        print(f"OK УстановитьСоединение({base}): {conn.Metadata.Synonym}")
        conn = None
    except Exception as e:
        print(f"FAIL УстановитьСоединение({base}): {err_text(e)}")
        fails += 1

# 4. СтрокаПодключения
s = M.СтрокаПодключения("bas_industrialbud")
if 'Srvr="localhost"' in s and 'Ref="bas_industrialbud"' in s:
    print(f"OK СтрокаПодключения: {s.replace('2442', '***')}")
else:
    print(f"FAIL СтрокаПодключения: {s.replace('2442', '***')}")
    fails += 1

# 5. Негатив: понятная ошибка без пароля
try:
    M.УстановитьСоединение("НЕТ_ТАКОЙ_БАЗЫ")
    print("FAIL негатив: исключение не брошено")
    fails += 1
except Exception as e:
    txt = err_text(e)
    if "2442" in txt:
        print("FAIL негатив: ПАРОЛЬ УТЁК В ТЕКСТ ОШИБКИ")
        fails += 1
    else:
        print(f"OK негатив: {txt.strip()[:160]}")

print("ИТОГ:", "OK" if fails == 0 else f"FAIL ({fails})")
