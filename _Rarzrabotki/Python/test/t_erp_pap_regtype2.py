import win32com.client, sys
sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch('V83.COMConnector')
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    ПАП.Регистратор КАК Рег,
    ПАП.ВидДвижения КАК ВД,
    ПАП.Сумма КАК Сумма,
    ПАП.Подразделение КАК Подр
ИЗ
    РегистрНакопления.ПрочиеАктивыПассивы КАК ПАП
ГДЕ
    ПАП.Период МЕЖДУ ДАТАВРЕМЯ(2025,12,1,0,0,0) И ДАТАВРЕМЯ(2026,6,30,23,59,59)
    И ТИПЗНАЧЕНИЯ(ПАП.Аналитика) = ТИП(Перечисление.ТипыНалогов)
"""
r = q.Execute().Выгрузить()
total = r.Количество()
# проверим уникальные значения ВД
vds=set()
for i in range(min(total,50)):
    vds.add(erp.XMLСтрока(r.Получить(i).ВД))
print("Уник.значения ВидДвижения (sample):", vds)

byreg = {}
for i in range(total):
    s = r.Получить(i)
    tp = s.Рег.Метаданные().Имя
    uid = erp.String(s.Рег.УникальныйИдентификатор())
    d = byreg.setdefault(tp, {"docs":set(),"prih":0.0,"rash":0.0})
    d["docs"].add(uid)
    vd = erp.XMLСтрока(s.ВД)
    if vd=="Receipt" or "риход" in vd:
        d["prih"]+=float(s.Сумма)
    else:
        d["rash"]+=float(s.Сумма)
print(f"\nВсего записей ПАП налоги: {total}")
for tp,d in sorted(byreg.items(), key=lambda x:-(x[1]['prih']+x[1]['rash'])):
    print(f"  {tp[:42]:<42} док={len(d['docs']):>3} Прих={d['prih']:>15,.2f} Расх={d['rash']:>15,.2f}")
