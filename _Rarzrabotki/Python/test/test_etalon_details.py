# -*- coding: utf-8 -*-
"""Detalnye dvizheniya etalona 000C-000001"""
import win32com.client
import pythoncom
from datetime import datetime

CONN_STR = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect(CONN_STR)
print("OK")

q = conn.NewObject("Query")
q.Text = """ВЫБРАТЬ Док.Ссылка ИЗ Документ.ВзаимозачетЗадолженности КАК Док
ГДЕ Док.Номер ПОДОБНО "%000Ц-000001%" И Док.Дата МЕЖДУ &Нач И &Кон"""
q.SetParameter("Нач", datetime(2025, 12, 1))
q.SetParameter("Кон", datetime(2025, 12, 31))
res = q.Execute()
sel = res.Choose()
sel.Next()
etalon = sel.Ссылка

# RaschetySPostavshchikami
print("\n=== RaschetySPostavshchikami ===")
q4 = conn.NewObject("Query")
q4.Text = """ВЫБРАТЬ Р.ВидДвижения, Р.ОбъектРасчетов, Р.Сумма, Р.СуммаУпр, Р.СуммаРегл
ИЗ РегистрНакопления.РасчетыСПоставщиками КАК Р ГДЕ Р.Регистратор = &Рег"""
q4.SetParameter("Рег", etalon)
res4 = q4.Execute()
sel4 = res4.Choose()
while sel4.Next():
    print(f"  Vid={sel4.ВидДвижения} Obj={sel4.ОбъектРасчетов} Summa={sel4.Сумма} SummaUpr={sel4.СуммаУпр} SummaRegl={sel4.СуммаРегл}")

# RaschetySKlientamiPoDocumentam
print("\n=== RaschetySKlientamiPoDocumentam ===")
q5 = conn.NewObject("Query")
q5.Text = """ВЫБРАТЬ Р.ВидДвижения, Р.СуммаУпр
ИЗ РегистрНакопления.РасчетыСКлиентамиПоДокументам КАК Р ГДЕ Р.Регистратор = &Рег"""
q5.SetParameter("Рег", etalon)
res5 = q5.Execute()
sel5 = res5.Choose()
while sel5.Next():
    print(f"  Vid={sel5.ВидДвижения} Obj={sel5.ОбъектРасчетов} Summa={sel5.Сумма}")

# ProchieAktivyPassivy cherez Oboroty
print("\n=== ProchieAktivyPassivy (Oboroty) ===")
q6 = conn.NewObject("Query")
q6.Text = """ВЫБРАТЬ
    Об.Подразделение, Об.Статья, Об.СуммаПриход, Об.СуммаРасход
ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Обороты(
    ДАТАВРЕМЯ(2025,12,17), ДАТАВРЕМЯ(2025,12,17,23,59,59), Регистратор,
) КАК Об
ГДЕ Об.Регистратор = &Рег"""
q6.SetParameter("Рег", etalon)
res6 = q6.Execute()
sel6 = res6.Choose()
while sel6.Next():
    dept = sel6.Подразделение.Наименование if conn.ЗначениеЗаполнено(sel6.Подразделение) else ""
    stat = sel6.Статья.Наименование if conn.ЗначениеЗаполнено(sel6.Статья) else ""
    print(f"  {dept} | {stat} | Db={float(sel6.СуммаПриход):.2f} | Kr={float(sel6.СуммаРасход):.2f}")

# Sravnim DebitorskayaZadolzhennost etalona s nashim
print("\n=== Etalon: DebitorskayaZadolzhennost detali ===")
q7 = conn.NewObject("Query")
q7.Text = """ВЫБРАТЬ
    ТЧ.Партнер,
    ТЧ.ТипРасчетов,
    ТЧ.ОбъектРасчетов,
    ТЧ.СуммаВзаиморасчетов,
    ТЧ.СуммаРегл,
    ТЧ.СуммаУпр,
    ТЧ.Сумма,
    ТЧ.ВалютаВзаиморасчетов,
    ТЧ.Организация,
    ТЧ.ИдентификаторСтроки
ИЗ Документ.ВзаимозачетЗадолженности.ДебиторскаяЗадолженность КАК ТЧ
ГДЕ ТЧ.Ссылка = &Рег"""
q7.SetParameter("Рег", etalon)
res7 = q7.Execute()
sel7 = res7.Choose()
while sel7.Next():
    print(f"  Partner={sel7.Партнер}")
    print(f"  TipRaschetov={sel7.ТипРасчетов}")
    print(f"  ObjRaschetov={sel7.ОбъектРасчетов}")
    print(f"  SummaVzaim={sel7.СуммаВзаиморасчетов}")
    print(f"  SummaRegl={sel7.СуммаРегл}")
    print(f"  SummaUpr={sel7.СуммаУпр}")
    print(f"  Summa={sel7.Сумма}")
    print(f"  Valuta={sel7.ВалютаВзаиморасчетов}")
    print(f"  Org={sel7.Организация}")
    print(f"  IdStr={sel7.ИдентификаторСтроки}")

print("\n=== Nash: DebitorskayaZadolzhennost ===")
q8 = conn.NewObject("Query")
q8.Text = """ВЫБРАТЬ Док.Ссылка ИЗ Документ.ВзаимозачетЗадолженности КАК Док
ГДЕ Док.А_ДокументОснованиеДляБаланса ССЫЛКА Документ.СписаниеНедостачТоваров
    И НЕ Док.ПометкаУдаления И Док.Дата МЕЖДУ &Нач И &Кон"""
q8.SetParameter("Нач", datetime(2025, 12, 1))
q8.SetParameter("Кон", datetime(2025, 12, 31, 23, 59, 59))
res8 = q8.Execute()
sel8 = res8.Choose()
if sel8.Next():
    nash = sel8.Ссылка
    q9 = conn.NewObject("Query")
    q9.Text = """ВЫБРАТЬ
        ТЧ.Партнер, ТЧ.ТипРасчетов, ТЧ.ОбъектРасчетов,
        ТЧ.СуммаВзаиморасчетов, ТЧ.СуммаРегл, ТЧ.СуммаУпр, ТЧ.Сумма,
        ТЧ.ВалютаВзаиморасчетов, ТЧ.Организация, ТЧ.ИдентификаторСтроки
    ИЗ Документ.ВзаимозачетЗадолженности.ДебиторскаяЗадолженность КАК ТЧ
    ГДЕ ТЧ.Ссылка = &Рег"""
    q9.SetParameter("Рег", nash)
    res9 = q9.Execute()
    sel9 = res9.Choose()
    while sel9.Next():
        print(f"  Partner={sel9.Партнер}")
        print(f"  TipRaschetov={sel9.ТипРасчетов}")
        print(f"  ObjRaschetov={sel9.ОбъектРасчетов}")
        print(f"  SummaVzaim={sel9.СуммаВзаиморасчетов}")
        print(f"  SummaRegl={sel9.СуммаРегл}")
        print(f"  SummaUpr={sel9.СуммаУпр}")
        print(f"  Summa={sel9.Сумма}")
        print(f"  Valuta={sel9.ВалютаВзаиморасчетов}")
        print(f"  Org={sel9.Организация}")
        print(f"  IdStr={sel9.ИдентификаторСтроки}")
