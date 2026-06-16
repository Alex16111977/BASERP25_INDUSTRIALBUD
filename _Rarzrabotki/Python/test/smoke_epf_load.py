# -*- coding: utf-8 -*-
# Smoke: внешняя обработка загружается через COM (целостность .epf)
import sys
sys.path.insert(0, r'C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test')
import _common_neizv as c

erp = c.connect()
path = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\ОбработкаНеизвестногоПартнера.epf"
try:
    obr = erp.ВнешниеОбработки.Создать(path, False)
    print("OK: обработка загружена через COM:", obr)
    print("Метаданные:", obr.Метаданные().Имя if hasattr(obr, 'Метаданные') else "?")
except Exception as e:
    info = getattr(e, 'excepinfo', None)
    print("FAIL:", info[2] if info else e)
