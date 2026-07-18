# -*- coding: utf-8 -*-
"""
Proverka: kakie registry zapolnyaet nash dokument vs etalon
"""
import win32com.client
import pythoncom
from datetime import datetime

CONN_STR = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect(CONN_STR)
print("OK")

# Najti nash dokument 000C-000002
q = conn.NewObject("Query")
q.Text = """ВЫБРАТЬ Док.Ссылка ИЗ Документ.ВзаимозачетЗадолженности КАК Док
ГДЕ Док.А_ДокументОснованиеДляБаланса ССЫЛКА Документ.СписаниеНедостачТоваров
    И НЕ Док.ПометкаУдаления
    И Док.Дата МЕЖДУ &Нач И &Кон"""
q.SetParameter("Нач", datetime(2025, 12, 1))
q.SetParameter("Кон", datetime(2025, 12, 31, 23, 59, 59))
res = q.Execute()
sel = res.Choose()
if not sel.Next():
    print("Ne najden!")
    exit()

dok_ref = sel.Ссылка
dok_obj = dok_ref.ПолучитьОбъект()
print(f"Dokument: {dok_ref}")
print(f"Proveden: {dok_obj.Проведен}")
print(f"TipDebitora: {dok_obj.ТипДебитора}")
print(f"TipKreditora: {dok_obj.ТипКредитора}")
print(f"VidOperacii: {dok_obj.ВидОперации}")
print(f"KontragentDebitor: {dok_obj.КонтрагентДебитор}")
print(f"KontragentKreditor: {dok_obj.КонтрагентКредитор}")
print(f"SummaRegl: {dok_obj.СуммаРегл}")

# Proverim dvizheniya
print("\n=== Dvizheniya ===")
registries = [
    "ПрочиеАктивыПассивы",
    "РасчетыСКлиентами",
    "РасчетыСПоставщиками",
    "РасчетыСКлиентамиПоДокументам",
    "РасчетыСПоставщикамиПоДокументам",
    "ПрочиеДоходы",
    "ПрочиеРасходы",
    "ДвиженияКонтрагентКонтрагент",
]
for reg_name in registries:
    try:
        q2 = conn.NewObject("Query")
        q2.Text = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол ИЗ РегистрНакопления.{reg_name} КАК Р ГДЕ Р.Регистратор = &Рег"
        q2.SetParameter("Рег", dok_ref)
        res2 = q2.Execute()
        sel2 = res2.Choose()
        sel2.Next()
        cnt = int(sel2.Кол)
        if cnt > 0:
            print(f"  {reg_name}: {cnt} zapisej")
    except:
        pass

# Teper' sravnim s etalonom 000C-000001
print("\n=== Etalon 000C-000001 ===")
q3 = conn.NewObject("Query")
q3.Text = """ВЫБРАТЬ Док.Ссылка ИЗ Документ.ВзаимозачетЗадолженности КАК Док
ГДЕ Док.Номер ПОДОБНО "%000Ц-000001%"
    И Док.Дата МЕЖДУ &Нач И &Кон"""
q3.SetParameter("Нач", datetime(2025, 12, 1))
q3.SetParameter("Кон", datetime(2025, 12, 31))
res3 = q3.Execute()
sel3 = res3.Choose()
if sel3.Next():
    etalon_ref = sel3.Ссылка
    etalon_obj = etalon_ref.ПолучитьОбъект()
    print(f"Etalon: {etalon_ref}")
    print(f"Proveden: {etalon_obj.Проведен}")
    print(f"TipDebitora: {etalon_obj.ТипДебитора}")
    print(f"TipKreditora: {etalon_obj.ТипКредитора}")

    for reg_name in registries:
        try:
            q4 = conn.NewObject("Query")
            q4.Text = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол ИЗ РегистрНакопления.{reg_name} КАК Р ГДЕ Р.Регистратор = &Рег"
            q4.SetParameter("Рег", etalon_ref)
            res4 = q4.Execute()
            sel4 = res4.Choose()
            sel4.Next()
            cnt = int(sel4.Кол)
            if cnt > 0:
                print(f"  {reg_name}: {cnt} zapisej")
        except:
            pass
