# -*- coding: utf-8 -*-
"""Подтверждение: ОтменитьТранзакцию-ветка записала ошибку в ЖурналРегистрации."""
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
import win32com.client
v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

ТЗ = buh.NewObject("ТаблицаЗначений")
Отбор = buh.NewObject("Структура")
Отбор.Вставить("Уровень", buh.УровеньЖурналаРегистрации.Ошибка)
Отбор.Вставить("ИмяСобытия", "СозданиеКомплектацийБух.Заповнення документів")
buh.ВыгрузитьЖурналРегистрации(ТЗ, Отбор)
print("записей в журнале по событию:", ТЗ.Количество())
if ТЗ.Количество() > 0:
    посл = ТЗ.Получить(ТЗ.Количество() - 1)
    print("последняя дата:", посл.Дата)
    print("комментарий:", (str(посл.Комментарий) or "")[:140])
    print("JOURNAL CHECK: PASS")
else:
    print("JOURNAL CHECK: записей нет (ветка отката не логировала?)")
