import win32com.client, sys
sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch('V83.COMConnector')
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

# Берём записи ПАП по налогам и определяем тип регистратора в Python
q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    ПАП.Регистратор КАК Рег,
    ПАП.Аналитика КАК Аналитика,
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
byreg = {}
for i in range(total):
    s = r.Получить(i)
    tp = s.Рег.Метаданные().Имя
    d = byreg.setdefault(tp, {"docs":set(),"prih":0.0,"rash":0.0,"podr_empty":0,"podr_fill":0})
    d["docs"].add(str(s.Рег.УникальныйИдентификатор()))
    if "Приход" in erp.XMLСтрока(s.ВД):
        d["prih"]+=float(s.Сумма)
    else:
        d["rash"]+=float(s.Сумма)
    if s.Подр.Пустая(): d["podr_empty"]+=1
    else: d["podr_fill"]+=1
print(f"Всего записей ПАП налоги: {total}")
print("=== По типу регистратора ===")
for tp,d in sorted(byreg.items(), key=lambda x:-(x[1]['prih']+x[1]['rash'])):
    print(f"  {tp[:42]:<42} док={len(d['docs']):>3} Прих={d['prih']:>15,.2f} Расх={d['rash']:>15,.2f} подрПусто={d['podr_empty']} подрЗап={d['podr_fill']}")
