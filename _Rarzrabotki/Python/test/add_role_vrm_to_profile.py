# -*- coding: utf-8 -*-
"""Добавить роль А_ДобавлениеИзменениеВыполненияРаботМодулей в профиль 'Менеджер производства'.

Идемпотентно: если роль уже в ТЧ Роли профиля — ничего не пишет.
Контроль: роль-маркер А_ВыполнениеРаботПросмотрЧасов в профиле отсутствует.
"""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

ROLE = "А_ДобавлениеИзменениеВыполненияРаботМодулей"
MARKER = "А_ВыполнениеРаботПросмотрЧасов"
PROFILE = "Менеджер производства"

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')


def find_role_id(name):
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ Иды.Ссылка КАК Ссылка
    ИЗ Справочник.ИдентификаторыОбъектовМетаданных КАК Иды
    ГДЕ Иды.ПолноеИмя = &ПолноеИмя
    """
    q.SetParameter("ПолноеИмя", "Роль." + name)
    r = q.Execute().Выгрузить()
    return r.Получить(0).Ссылка if r.Количество() > 0 else None


role_id = find_role_id(ROLE)
if role_id is None:
    # БСП сам регистрирует отсутствующий идентификатор при обращении
    print("Идентификатор роли не найден — регистрирую через ОбщегоНазначения.ИдентификаторОбъектаМетаданных...")
    role_id = erp.ОбщегоНазначения.ИдентификаторОбъектаМетаданных("Роль." + ROLE)
if role_id is None:
    print("FAIL: идентификатор роли так и не найден")
    sys.exit(1)
print("OK: идентификатор роли найден")

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ Профили.Ссылка КАК Ссылка
ИЗ Справочник.ПрофилиГруппДоступа КАК Профили
ГДЕ Профили.Наименование = &Наим И НЕ Профили.ПометкаУдаления
"""
q.SetParameter("Наим", PROFILE)
r = q.Execute().Выгрузить()
if r.Количество() == 0:
    print("FAIL: профиль не найден")
    sys.exit(1)
prof_ref = r.Получить(0).Ссылка
prof_obj = prof_ref.ПолучитьОбъект()

names = []
for i in range(prof_obj.Роли.Количество()):
    stroka = prof_obj.Роли.Получить(i)
    try:
        names.append(stroka.Роль.ПолноеИмя)
    except Exception:
        pass

if ("Роль." + MARKER) in names:
    print(f"FAIL: в профиле обнаружена роль-маркер {MARKER} — колонки часов будут видны!")
    sys.exit(1)

if ("Роль." + ROLE) in names:
    print("OK: роль уже в профиле, запись не требуется")
else:
    nr = prof_obj.Роли.Добавить()
    nr.Роль = role_id
    prof_obj.Записать()
    print("OK: роль добавлена в профиль, профиль записан (БСП пересчитает права групп)")

# контрольное перечитывание
prof_obj2 = prof_ref.ПолучитьОбъект()
names2 = []
for i in range(prof_obj2.Роли.Количество()):
    try:
        names2.append(prof_obj2.Роли.Получить(i).Роль.ПолноеИмя)
    except Exception:
        pass
print("Контроль: роль в ТЧ =", ("Роль." + ROLE) in names2,
      "; маркер отсутствует =", ("Роль." + MARKER) not in names2,
      "; всего ролей =", len(names2))
print("DONE")
