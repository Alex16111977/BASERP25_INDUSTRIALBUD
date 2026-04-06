# -*- coding: utf-8 -*-
"""
Test: "zayvie" vzaimozachety pri filtre po podrazdeleniyu.
Problem: IB00-001591 (Globino-2/KMD-2) popadaet kak "zayviy" pri filtre Astarta
potomu chto on ne v spiske disbalansov (filtr isklyuchil).
Reshenie: v bloke zayvikh tozhe primenyat filtr po podrazdeleniyu.
"""
import win32com.client, pythoncom
from datetime import datetime
from collections import defaultdict

CONN_STR = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect(CONN_STR)
print("OK\n")

NACH = datetime(2025, 12, 1)
KON = datetime(2025, 12, 31, 23, 59, 59)
FILTR_DEPT = "Астарта. Тищенки"

# Najti VSE vzaimozachety s A_DokumentOsnovaniyeDlyaBalansa za period
q = conn.NewObject("Query")
q.Text = """ВЫБРАТЬ
    Док.Ссылка КАК Ссылка,
    Док.А_ДокументОснованиеДляБаланса КАК ДокОснование,
    Док.СуммаРегл КАК Сумма,
    Док.КонтрагентДебитор КАК Дебитор,
    Док.КонтрагентКредитор КАК Кредитор
ИЗ Документ.ВзаимозачетЗадолженности КАК Док
ГДЕ Док.Дата МЕЖДУ &Нач И &Кон
    И Док.А_ДокументОснованиеДляБаланса <> НЕОПРЕДЕЛЕНО
    И Док.Проведен
    И НЕ Док.ПометкаУдаления"""
q.SetParameter("Нач", NACH)
q.SetParameter("Кон", KON)

sel = q.Execute().Choose()
print("=== Vzaimozachety s DokOsnovanie ===")
while sel.Next():
    deb = sel.Дебитор.Наименование if conn.ЗначениеЗаполнено(sel.Дебитор) else "?"
    kred = sel.Кредитор.Наименование if conn.ЗначениеЗаполнено(sel.Кредитор) else "?"
    dok_osn = str(sel.ДокОснование) if conn.ЗначениеЗаполнено(sel.ДокОснование) else "?"
    print(f"  {str(sel.Ссылка):50s}")
    print(f"    DokOsn: {dok_osn}")
    print(f"    Deb={deb:25s} Kred={kred:25s} Sum={float(sel.Сумма):.2f}")

    # Proverka: dolzhen li etot vzaimozachet popadat pri filtre Astarta?
    has_astarta = (FILTR_DEPT in deb or FILTR_DEPT in kred)
    print(f"    Filtr '{FILTR_DEPT}': {'DA - popadaet' if has_astarta else 'NET - NE dolzhen popadat'}")
    print()

print("\n=== VYVOD ===")
print(f"Pri filtre '{FILTR_DEPT}' v blok 'zayvikh' dolzhny popadat")
print("tolko vzaimozachety gde debitor ILI kreditor = '{FILTR_DEPT}'.")
print("Ostalnye - propuskat (oni otnosyatsya k drugim podrazdeleniyam).")
