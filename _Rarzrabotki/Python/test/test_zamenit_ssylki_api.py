# -*- coding: utf-8 -*-
# Task 3: подтвердить структуру параметров ОбщегоНазначения.ПараметрыЗаменыСсылок
import sys
sys.path.insert(0, r'C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test')
import _common_neizv as c

erp = c.connect()

try:
    prm = erp.ОбщегоНазначения.ПараметрыЗаменыСсылок()
    print("Тип:", type(prm))
    for kv in prm:
        print("  param:", kv.Ключ, "=", repr(kv.Значение))
except Exception as e:
    info = getattr(e, 'excepinfo', None)
    print("FAIL ПараметрыЗаменыСсылок:", info[2] if info else e)

# Проверим сигнатуру ЗаменитьСсылки наличием метода
print("\nЕсть метод ЗаменитьСсылки:", hasattr(erp.ОбщегоНазначения, "ЗаменитьСсылки"))
