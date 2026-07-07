# -*- coding: utf-8 -*-
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

# Поля основной таблицы Хозрасчетного через пустой запрос-структуру не вытащить просто,
# попробуем СубконтоКт1 (без Дт-варианта) и виртуальную ДвиженияССубконто
qs = buh.NewObject("Запрос")
qs.Text = """ВЫБРАТЬ Сч.Ссылка КАК Ссылка ИЗ ПланСчетов.Хозрасчетный КАК Сч ГДЕ Сч.Код = "6853" """
счет = qs.Execute().Выгрузить().Получить(0).Ссылка

# Через виртуальную ДвиженияССубконто
q = buh.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ РАЗЛИЧНЫЕ
    ВЫБОР КОГДА Дв.СчетКт = &Счет ТОГДА Дв.СубконтоКт1 ИНАЧЕ Дв.СубконтоДт1 КОНЕЦ КАК Суб
ИЗ РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто КАК Дв
ГДЕ (Дв.СчетДт = &Счет ИЛИ Дв.СчетКт = &Счет)
"""
q.SetParameter("Счет", счет)
try:
    r = q.Execute().Выгрузить()
    types={}; samples={}
    for i in range(r.Количество()):
        s=r.Получить(i)
        if buh.ЗначениеЗаполнено(s.Суб):
            tn=s.Суб.Метаданные().Имя
            types[tn]=types.get(tn,0)+1
            if tn not in samples: samples[tn]=str(s.Суб)
        else:
            types["Пусто"]=types.get("Пусто",0)+1
    print("Типы субконто получателя (ДвиженияССубконто):", types)
    print("Примеры:", samples)
except Exception as e:
    if hasattr(e,'excepinfo') and e.excepinfo:
        print("FAIL:", e.excepinfo[2])
    else:
        print("FAIL:", e)
