import win32com.client, sys, json
sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
c = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
FILTR = """ (Д.Подразделение.Наименование ПОДОБНО "%IRC%" ИЛИ Д.Подразделение.Наименование ПОДОБНО "%IRС%")
    И НЕ Д.ПометкаУдаления """
q = c.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Стр.Ссылка.Номер КАК Номер,
    ПРЕДСТАВЛЕНИЕ(Стр.ОбщееНазвание) КАК ОН,
    Стр.Номенклатура.Наименование КАК НомНаим,
    ПРЕДСТАВЛЕНИЕ(Стр.Единица) КАК Ед,
    СУММА(Стр.Количество) КАК Кол,
    СУММА(Стр.Сумма) КАК Сумма
ИЗ Документ.РасчетКомплектаций.СписаниеПоНормам КАК Стр
ГДЕ Стр.Ссылка В (ВЫБРАТЬ Д.Ссылка ИЗ Документ.РасчетКомплектаций КАК Д ГДЕ """ + FILTR + """)
СГРУППИРОВАТЬ ПО Стр.Ссылка.Номер, ПРЕДСТАВЛЕНИЕ(Стр.ОбщееНазвание), Стр.Номенклатура.Наименование, ПРЕДСТАВЛЕНИЕ(Стр.Единица)"""
try:
    r = q.Execute().Выгрузить()
except Exception as e:
    print('FAIL', e.excepinfo[2] if getattr(e,'excepinfo',None) else e); sys.exit(1)
out=[]
for i in range(r.Количество()):
    x=r.Получить(i)
    out.append(dict(док=x.Номер, он=str(x.ОН), ном=str(x.НомНаим or ''), ед=str(x.Ед), кол=float(x.Кол or 0), сум=float(x.Сумма or 0)))
json.dump(out, open('pn.json','w',encoding='utf-8'), ensure_ascii=False)
print('OK', len(out))
