# -*- coding: utf-8 -*-
"""Acceptance переноса регистра А_ОтчетБаланс_Свод -> OlapBASERP.Fact_Balance.

Текущий оркестратор А_ФинРез_Баланс проводит ТОЛЬКО Свод_СебестоимостьТоваров
(остальные Свод_* закомментированы), поэтому полный критерий Актив=Пасив
(Σ Sum_Close≈0) НЕприменим — он для ПОЛНОГО баланса. Здесь проверяем именно
КОРРЕКТНОСТЬ ОБМЕНА:
 1) Fact_Balance(OLAP) 1:1 отражает РегистрСведений.А_ОтчетБаланс_Свод (ETL
    только копирует): COUNT и Σ 4 ресурсов совпадают;
 2) Σ Sum_Close по PAP_Article == ПАП.ОстаткиИОбороты (январь/ТОВ, OD-3) до
    копейки (логика штатного test_etl_acceptance_balance.py, проверка #2);
 3) «Товары на оптовых складах» Sum_Close == эталон 83 627 719,44.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pyodbc, pythoncom, win32com.client

OLAP = ("DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;"
        "DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;")
ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
TOL = 0.01
ЭТАЛОН = 83_627_719.44

cx = pyodbc.connect(OLAP)
# (1) агрегаты Fact_Balance за январь
fb = cx.execute("""SELECT COUNT(*), ISNULL(SUM(Sum_Open),0),
        ISNULL(SUM(Sum_Inflow),0), ISNULL(SUM(Sum_Outflow),0),
        ISNULL(SUM(Sum_Close),0)
    FROM Fact_Balance WHERE Period_Month='2026-01-01'""").fetchone()
fb_cnt, fb_o, fb_i, fb_ou, fb_c = (int(fb[0]), float(fb[1]), float(fb[2]),
                                   float(fb[3]), float(fb[4]))
olap_st = {r[0]: float(r[1] or 0) for r in cx.execute(
    """SELECT d.PAP_Article_Name, SUM(f.Sum_Close)
       FROM Fact_Balance f JOIN Dim_PAP_Articles d
         ON f.PAP_Article_ID = d.PAP_Article_ID
       WHERE f.Period_Month='2026-01-01'
       GROUP BY d.PAP_Article_Name""").fetchall()}
cx.close()
print(f"Fact_Balance(2026-01): rows={fb_cnt} ΣOpen={fb_o:,.2f} "
      f"ΣInflow={fb_i:,.2f} ΣOutflow={fb_ou:,.2f} ΣClose={fb_c:,.2f}")

pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(ERP)

# (1) тот же агрегат из регистра 1С — ETL обязан совпасть 1:1
qr = conn.NewObject("Запрос")
qr.Текст = """ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К,
    СУММА(СуммаНачальныйОстаток) КАК О, СУММА(СуммаПриход) КАК П,
    СУММА(СуммаРасход) КАК Р, СУММА(СуммаКонечныйОстаток) КАК З
ИЗ РегистрСведений.А_ОтчетБаланс_Свод"""
rr = qr.Выполнить().Выбрать(); rr.Следующий()
rg_cnt = int(rr.К); rg_o = float(rr.О or 0); rg_i = float(rr.П or 0)
rg_ou = float(rr.Р or 0); rg_c = float(rr.З or 0)
print(f"Регистр А_ОтчетБаланс_Свод: rows={rg_cnt} ΣOpen={rg_o:,.2f} "
      f"ΣInflow={rg_i:,.2f} ΣOutflow={rg_ou:,.2f} ΣClose={rg_c:,.2f}")

# (2) эталон ПАП по статье (логика штатного теста, январь/ТОВ, OD-3)
q0 = conn.NewObject("Запрос")
q0.Текст = ('ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ Справочник.Организации '
            'ГДЕ КодПоЕДРПОУ = "40645273"')
s = q0.Выполнить().Выбрать(); s.Следующий(); org = s.С
qb = conn.NewObject("Запрос")
qb.Текст = ("ВЫБРАТЬ НАЧАЛОПЕРИОДА(ДАТАВРЕМЯ(2026,1,1),МЕСЯЦ) КАК НМ, "
            "КОНЕЦПЕРИОДА(ДАТАВРЕМЯ(2026,1,31,23,59,59),ДЕНЬ) КАК КД")
rb = qb.Выполнить().Выбрать(); rb.Следующий()
qi = conn.NewObject("Запрос")
qi.Текст = ('ВЫБРАТЬ Ссылка КАК С ИЗ '
            'ПланВидовХарактеристик.СтатьиАктивовПассивов ГДЕ Наименование В (&Сп)')
a = conn.NewObject("Массив")
for n in ["Собственные средства", "Доходы текущего периода",
          "Расходы текущего периода"]:
    a.Добавить(n)
qi.УстановитьПараметр("Сп", a)
искл = qi.Выполнить().Выгрузить().ВыгрузитьКолонку("С")
qp = conn.NewObject("Запрос")
qp.Текст = (
    "ВЫБРАТЬ ПРЕДСТАВЛЕНИЕ(Б.Статья) КАК Н, "
    "СУММА(Б.СуммаКонечныйОстаток) КАК sK "
    "ИЗ РегистрНакопления.ПрочиеАктивыПассивы.ОстаткиИОбороты(&Д1,&Д2,Авто,,"
    "Организация=&Орг И НЕ Статья В ИЕРАРХИИ(&Искл)) КАК Б "
    "СГРУППИРОВАТЬ ПО ПРЕДСТАВЛЕНИЕ(Б.Статья) "
    "ИМЕЮЩИЕ СУММА(Б.СуммаКонечныйОстаток) <> 0")
qp.УстановитьПараметр("Д1", rb.НМ); qp.УстановитьПараметр("Д2", rb.КД)
qp.УстановитьПараметр("Орг", org); qp.УстановитьПараметр("Искл", искл)
t = qp.Выполнить().Выгрузить()
pap = {str(t.Получить(i).Н): float(t.Получить(i).sK or 0)
       for i in range(t.Количество())}

# сверка ТОЛЬКО по статьям, которые реально есть в своде (себестоимость):
плохо = []
for k in olap_st:
    d = abs(olap_st[k] - pap.get(k, 0.0))
    if d > TOL:
        плохо.append((k, pap.get(k, 0.0), olap_st[k], d))
тов = olap_st.get("Товары на оптовых складах", 0.0)
print(f"\nстатей в Fact_Balance: {len(olap_st)}")
print(f"«Товары на оптовых складах» Sum_Close = {тов:,.2f} "
      f"| ПАП = {pap.get('Товары на оптовых складах',0.0):,.2f} "
      f"| эталон = {ЭТАЛОН:,.2f}")
for k, p, o, dl in плохо[:10]:
    print(f"  BAD {k[:34]:34} ПАП={p:,.2f} OLAP={o:,.2f} Δ={dl:,.4f}")

print("=" * 64)
assert fb_cnt == rg_cnt, f"FAIL: rows OLAP {fb_cnt} != регистр {rg_cnt}"
assert abs(fb_o - rg_o) < TOL and abs(fb_i - rg_i) < TOL \
    and abs(fb_ou - rg_ou) < TOL and abs(fb_c - rg_c) < TOL, \
    "FAIL: Σ ресурсов Fact_Balance != регистр (ETL не 1:1)"
assert not плохо, f"FAIL: Fact_Balance != ПАП по статьям ({len(плохо)})"
assert abs(тов - ЭТАЛОН) < TOL, \
    f"FAIL: «Товары» {тов:,.2f} != эталон {ЭТАЛОН:,.2f}"
print("PASS: обмен корректен — Fact_Balance 1:1 == регистр А_ОтчетБаланс_Свод; "
      "Σ Sum_Close по статьям == ПАП.ОстаткиИОбороты до копейки; "
      "«Товары на оптовых складах» == 83 627 719,44 (эталон ПАП/Отчёт)")
