# -*- coding: utf-8 -*-
"""Проверка появления Документ.РасчетКомплектаций в живой базе после db-update."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

md = buh.Метаданные.Документы.Найти("РасчетКомплектаций")
assert md is not None, "Документ.РасчетКомплектаций не найден"
en = buh.Метаданные.Перечисления.Найти("СтатусыРасчетаКомплектаций")
assert en is not None, "Перечисление не найдено"

tabs = [md.ТабличныеЧасти.Получить(i).Имя for i in range(md.ТабличныеЧасти.Количество())]
print("ТЧ:", ", ".join(tabs))
assert len(tabs) == 9, f"ожидал 9 ТЧ, есть {len(tabs)}"
for need in ("ТабличнаяЧастьОстатков", "СписаниеПоНормам", "СписаниеСверхНормы",
             "ДокументиКомплектації", "ДокументиМалоценки", "СчетаМалоценки"):
    assert need in tabs, f"нет ТЧ {need}"

makety = [md.Макеты.Получить(i).Имя for i in range(md.Макеты.Количество())]
print("Макеты:", ", ".join(makety))
assert sorted(makety) == sorted(["МакетПланФакт", "МакетПланФактЕтапи", "МакетАнализСС", "МакетАнализССОдна"]), makety

forms = [md.Формы.Получить(i).Имя for i in range(md.Формы.Количество())]
print("Формы:", ", ".join(forms))
assert "ФормаДокумента" in forms

attrs = [md.Реквизиты.Получить(i).Имя for i in range(md.Реквизиты.Количество())]
print("Реквизиты:", ", ".join(attrs))
assert len(attrs) == 7

vals = [en.ЗначенияПеречисления.Получить(i).Имя for i in range(en.ЗначенияПеречисления.Количество())]
print("Статусы:", ", ".join(vals))
assert vals == ["Черновик", "РасчетВыполнен", "ДокументыСозданы"]

print("CHECK LOADED PASS")
