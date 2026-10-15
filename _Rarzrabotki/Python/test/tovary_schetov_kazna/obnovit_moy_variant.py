# -*- coding: utf-8 -*-
"""Обновление пользовательского варианта «Прогноз за потребою — детально - моя» (Справочник.ВариантыОтчетов):
в отбор полей после «ЭкономияСС» добавляются ОплаченоНеНадійшло, Оплачено, Борг (формат ЧДЦ=2).
Остальное (структура, отбор подразделения, дата, оформление) не меняется. Резервная копия — рядом с файлом варианта.

    python obnovit_moy_variant.py          # пробный: показать, что изменится
    python obnovit_moy_variant.py --write  # записать
"""
import datetime
import re
import sys

import win32com.client

sys.stdout.reconfigure(encoding='utf-8')
WRITE = "--write" in sys.argv
KEY = "8bf5b43a-5bc3-43c0-9c0a-584169e689c4"
BACKUP_DIR = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\temp"
NEW = ["ОплаченоНеНадійшло", "Оплачено", "Борг"]
NS = "http://v8.1c.ru/8.1/data-composition-system/settings"

erp = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
z = erp.NewObject("Запрос")
z.Text = """ВЫБРАТЬ Вар.Ссылка ИЗ Справочник.ВариантыОтчетов КАК Вар
ГДЕ Вар.КлючВарианта = &К И Вар.Пользовательский И НЕ Вар.ПометкаУдаления"""
z.SetParameter("К", KEY)
r = z.Execute().Выгрузить()
assert r.Количество() == 1, "вариант не найден или не один"
obj = r.Получить(0).Получить(0).ПолучитьОбъект()
nastr = obj.Настройки.Получить()


def to_xml(settings):
    w = erp.NewObject("ЗаписьXML")
    w.УстановитьСтроку()
    erp.ФабрикаXDTO.ЗаписатьXML(w, erp.СериализаторXDTO.ЗаписатьXDTO(settings), "Settings", NS)
    return w.Закрыть()


xml = to_xml(nastr)
stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
with open(rf"{BACKUP_DIR}\Потреби мой вариант.backup_{stamp}.xml", "w", encoding="utf-8") as fh:
    fh.write(xml)
present = [f for f in NEW if f"<field>{f}</field>" in xml.split("</selection>")[0]]
if present:
    print("уже есть в отборе:", present)
    sys.exit(0)

# 1) поля в отбор — после ЭкономияСС (через API компоновки; текстовая правка XML ломает пространства имён)
nov = nastr
sel = nov.Выбор.Элементы
idx = next(i for i in range(sel.Количество())
           if hasattr(sel.Получить(i), "Поле") and erp.String(sel.Получить(i).Поле) == "ЭкономияСС")
for k, f in enumerate(NEW, 1):
    el = sel.Вставить(idx + k, erp.NewObject("ОписаниеТипов", "ВыбранноеПолеКомпоновкиДанных").Типы().Получить(0))
    el.Поле = erp.NewObject("ПолеКомпоновкиДанных", f)
    el.Использование = True

# 2) формат ЧДЦ=2 для новых полей
ca = nov.УсловноеОформление.Элементы.Добавить()
for f in NEW:
    ca.Поля.Элементы.Добавить().Поле = erp.NewObject("ПолеКомпоновкиДанных", f)
ca.Оформление.УстановитьЗначениеПараметра("Формат", "ЧДЦ=2")

names = [erp.String(sel.Получить(i).Поле) for i in range(sel.Количество()) if hasattr(sel.Получить(i), "Поле")]
print("отбор полей после правки:", names)
assert all(f in names for f in NEW)
if WRITE:
    obj.Настройки = erp.NewObject("ХранилищеЗначения", nov)
    obj.Записать()
    print("записано")
else:
    print("пробный режим: не записано")
