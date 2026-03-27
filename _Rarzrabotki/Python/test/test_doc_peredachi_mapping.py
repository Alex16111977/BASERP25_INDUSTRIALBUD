# -*- coding: utf-8 -*-
"""
Test: ДокументПередачи mapping from Kazna to ERP
Verifies that V83.XMLСтрока() works for COM UUID serialization
and that Kazna document UUIDs can be found in ERP.
"""
import win32com.client

v8 = win32com.client.Dispatch("V83.COMConnector")

# Connect to Kazna
CONN_KAZN = 'Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"'
print("Connecting to Kazna...")
kazna = v8.Connect(CONN_KAZN)

# Connect to ERP
CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
print("Connecting to ERP...")
erp = v8.Connect(CONN_ERP)

print("\n=== Step 1: Query Kazna register ДенежныеСредстваФинАгенты.ОстаткиИОбороты ===")

q = kazna.NewObject("Query")
q.Text = (
    "SELECT TOP 10 "
    "  DS.FinAgent AS FinAgent, "
    "  DS.Napravlenie AS Napravlenie, "
    "  DS.DokumentPeredachi AS DokPeredachi, "
    "  DS.SummaKonechnyyOstatok AS SumKonOst "
    "FROM "
    "  AccumulationRegister.DenezhnyyeSredstvaFinAgenty.BalanceAndTurnovers AS DS "
)

# Actually in 1C queries must be in Russian:
q.Text = (
    "ВЫБРАТЬ ПЕРВЫЕ 10 "
    "ДС.ФинАгент КАК ФинАгент, "
    "ДС.Направление КАК Направление, "
    "ДС.ДокументПередачи КАК ДокументПередачи, "
    "ДС.СуммаКонечныйОстаток КАК СуммаКонОст "
    "ИЗ РегистрНакопления.ДенежныеСредстваФинАгенты.ОстаткиИОбороты КАК ДС "
    "ГДЕ ДС.СуммаКонечныйОстаток <> 0"
)

print("Executing query...")
res = q.Execute()
sel = res.Choose()

found_docs = 0
empty_docs = 0

while sel.Next():
    fa = kazna.String(sel.ФинАгент)
    napr = ""
    try:
        napr = kazna.String(sel.Направление)
    except:
        pass

    dok = sel.ДокументПередачи
    sum_val = sel.СуммаКонОст

    is_filled = kazna.ЗначениеЗаполнено(dok)

    if is_filled:
        found_docs += 1

        # Test 1: V83.XMLСтрока (should work)
        try:
            uid_xml = kazna.XMLСтрока(dok.УникальныйИдентификатор())
            print(f"  OK V83.XMLСтрока: ФинАгент={fa}, Направление={napr}, UUID={uid_xml}, Сумма={sum_val}")
        except Exception as e:
            print(f"  FAIL V83.XMLСтрока: {e}")
            uid_xml = None

        # Test 2: V83.String (alternative approach)
        try:
            uid_str = kazna.String(dok.УникальныйИдентификатор())
            print(f"  OK V83.String:    UUID={uid_str}")
        except Exception as e:
            print(f"  FAIL V83.String: {e}")
            uid_str = None

        # Test 3: Direct XMLСтрока (this is what the broken code does)
        try:
            uid_direct = v8.XMLСтрока(dok.УникальныйИдентификатор())
            print(f"  OK direct XMLСтрока: UUID={uid_direct}")
        except Exception as e:
            print(f"  FAIL direct XMLСтрока (expected!): {e}")

        # Test 4: Check if UUID exists in ERP as СписаниеБезналичныхДенежныхСредств
        if uid_xml:
            try:
                erp_q = erp.NewObject("Query")
                erp_q.Text = (
                    "ВЫБРАТЬ ПЕРВЫЕ 1 "
                    "Д.Ссылка КАК Ссылка, "
                    "Д.Номер КАК Номер, "
                    "Д.Дата КАК Дата "
                    "ИЗ Документ.СписаниеБезналичныхДенежныхСредств КАК Д "
                    "ГДЕ Д.Ссылка = &Ссылка"
                )
                erp_ref = erp.Документы.СписаниеБезналичныхДенежныхСредств.ПолучитьСсылку(
                    erp.NewObject("УникальныйИдентификатор", uid_xml))
                erp_q.SetParameter("Ссылка", erp_ref)
                erp_res = erp_q.Execute()
                if not erp_res.Empty():
                    erp_sel = erp_res.Choose()
                    erp_sel.Next()
                    print(f"  ERP FOUND: {erp.String(erp_sel.Ссылка)}")
                else:
                    print(f"  ERP NOT FOUND for UUID {uid_xml}")
            except Exception as e:
                print(f"  ERP check error: {e}")

        print()
    else:
        empty_docs += 1

print(f"\n=== Summary ===")
print(f"Rows with filled ДокументПередачи: {found_docs}")
print(f"Rows with empty ДокументПередачи: {empty_docs}")
print("Done!")
