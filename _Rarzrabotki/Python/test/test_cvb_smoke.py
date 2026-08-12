# -*- coding: utf-8 -*-
"""Smoke А_ЦентрВыравниванияБаз: компиляция, seed реестра, план Казны из метаданных,
round-trip хранилища настроек, узлы обмена. Только чтение (кроме служебного ключа настроек)."""
import sys
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\А_ЦентрВыравниванияБаз.epf"
PLAN_BUH = "ОбменУправлениеПредприятиемБухгалтерия20"

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

obj = erp.ВнешниеОбработки.Создать(EPF, False)
print("1. Создание/компиляция ObjectModule: OK")

obj.ИнициализироватьРеестрПоУмолчанию()
n = obj.РеестрКонтуров.Количество()
assert n == 4, "реестр: ожидалось 4, получено %s" % n
names = [str(obj.РеестрКонтуров.Получить(i).ИмяКонтура) for i in range(n)]
plans = [str(obj.РеестрКонтуров.Получить(i).ИмяПланаОбмена) for i in range(n)]
paths = [str(obj.РеестрКонтуров.Получить(i).ПутьКОбработке) for i in range(n)]
print("2. Реестр:", names)

kazna_meta = None
mp = erp.Metadata.ExchangePlans
for i in range(mp.Count()):
    nm = str(mp.Get(i).Name)
    if nm.startswith("Казначе"):
        kazna_meta = nm
assert kazna_meta, "план Казны не найден в метаданных"
assert plans[3] == kazna_meta, "план Казны в реестре (%r) != метаданные (%r)" % (plans[3], kazna_meta)
assert plans[0] == PLAN_BUH and plans[1] == PLAN_BUH and plans[2] == PLAN_BUH
print("3. Планы обмена реестра корректны (Казна:", kazna_meta, ")")

import os
for p in paths:
    assert os.path.exists(p), "нет файла плагина: %s" % p
print("4. Все 4 файла плагинов существуют")

obj.СохранитьНастройки()
obj2 = erp.ВнешниеОбработки.Создать(EPF, False)
was = obj2.ЗагрузитьНастройки()
assert was, "ЗагрузитьНастройки вернул Ложь после сохранения"
assert obj2.РеестрКонтуров.Количество() == 4
print("5. Хранилище настроек round-trip: OK")

u1 = obj.УзелКорреспондентПлана(PLAN_BUH)
u2 = obj.УзелКорреспондентПлана(kazna_meta)
assert u1 is not None and u2 is not None, "узлы обмена не найдены"
print("6. Узлы обмена:", erp.String(u1), "|", erp.String(u2))

grp = obj.ИменаПлановГрупп()
assert grp.Количество() == 2, "ожидалось 2 группы планов"
assert str(grp.Получить(0)) == PLAN_BUH, "первая группа должна быть Бух (порядок 10)"
idx_buh = obj.ИндексыКонтуровПлана(PLAN_BUH)
idx_kaz = obj.ИндексыКонтуровПлана(kazna_meta)
assert idx_buh.Количество() == 3 and idx_kaz.Количество() == 1
print("7. Группировка по планам: Бух=3 контура, Казна=1")

print("SMOKE OK")
