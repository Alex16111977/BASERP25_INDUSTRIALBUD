# -*- coding: utf-8 -*-
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
def q(t):
    x = erp.NewObject("Запрос"); x.Text = t; return x.Execute().Выгрузить()

n = q("ВЫБРАТЬ КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ПР.Регистратор) КАК К ИЗ РегистрНакопления.ПрочиеРасходы КАК ПР").Получить(0).К
print("Уникальных регистраторов в ПрочиеРасходы:", n)
s = q("ВЫБРАТЬ ПЕРВЫЕ 1 ПР.Регистратор КАК Р ИЗ РегистрНакопления.ПрочиеРасходы КАК ПР").Получить(0).Р
try:
    print("Тест .Метаданные().Имя на ссылке:", s.Метаданные().Имя)
except Exception as e:
    info = getattr(e, 'excepinfo', None); print("FAIL .Метаданные():", info[2] if info else e)
