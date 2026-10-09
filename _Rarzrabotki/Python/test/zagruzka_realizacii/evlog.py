# -*- coding: utf-8 -*-
"""Помилки з журналу реєстрації BuhBud за останні N хвилин (стек з модулем і рядком) — для помилок форми зовнішньої
обробки, яку збірка не компілює, а COM не відкриває.
Запуск: python evlog.py [хвилин=15]"""
import datetime
import sys

import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
ХВИЛИН = int(sys.argv[1]) if len(sys.argv) > 1 else 15
c = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
S = c.String
тз = c.NewObject("ТаблицаЗначений")
відбір = c.NewObject("Структура")
# дата через COM зсувається на -3 год — компенсуємо
відбір.Вставить("ДатаНачала", datetime.datetime.now() - datetime.timedelta(minutes=ХВИЛИН) + datetime.timedelta(hours=3))
рівні = c.NewObject("Массив")
рівні.Добавить(c.УровеньЖурналаРегистрации.Ошибка)
відбір.Вставить("Уровень", рівні)
c.UnloadEventLog(тз, відбір)  # англійський аліас ВыгрузитьЖурналРегистрации (кирилиця через COM не резолвиться)
print("помилок:", тз.Количество())
for i in range(тз.Количество()):
    р = тз.Получить(i)
    print("----", S(р.Дата), S(р.ИмяПользователя), S(р.ПредставлениеСобытия))
    print(S(р.Комментарий)[:3000])
