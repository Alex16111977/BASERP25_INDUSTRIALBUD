# -*- coding: utf-8 -*-
"""Диагностика расхождения «Оплата труда»: 1С регистр vs ПАП vs SQL Fact_Balance (PowerBI).

Слои:
  A. РегистрСведений.А_ОтчетБаланс_Свод (что проведено в 1С, Статья=Оплата труда)
  B. РегистрНакопления.ПрочиеАктивыПассивы (источник истины, Статья=ОТ, Источник=ПустаяСсылка)
  C. OlapBASERP.Fact_Balance (то, что читает PowerBI PL.pbix)
"""
import sys, io, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

def f(x):
    try: return float(x)
    except: return 0.0

# Период исследования: дек 2025 .. апр 2026
month_ends = [
    ("2025-12", datetime.datetime(2025,12,31,23,59,59)),
    ("2026-01", datetime.datetime(2026, 1,31,23,59,59)),
    ("2026-02", datetime.datetime(2026, 2,28,23,59,59)),
    ("2026-03", datetime.datetime(2026, 3,31,23,59,59)),
    ("2026-04", datetime.datetime(2026, 4,30,23,59,59)),
]

# === Найти статью «Оплата труда» ===
q = erp.NewObject("Запрос")
q.Текст = ("ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка КАК С ИЗ ПланВидовХарактеристик.СтатьиАктивовПассивов "
           "ГДЕ Наименование = \"Оплата труда\"")
ОТ = q.Выполнить().Выгрузить()[0].С
print(f"Статья ОТ UUID = {erp.String(ОТ.УникальныйИдентификатор())}")

# === A. Регистр А_ОтчетБаланс_Свод — per месяц регистратора, по Source и Расхождение ===
print("\n=== A. РегистрСведений.А_ОтчетБаланс_Свод (Статья=Оплата труда) ===")
qa = erp.NewObject("Запрос")
qa.УстановитьПараметр("ОТ", ОТ)
qa.Текст = """
ВЫБРАТЬ
    Рег.Регистратор.Месяц КАК Месяц,
    Рег.Регистратор.Номер КАК Номер,
    Рег.Регистратор.Проведен КАК Проведен,
    ПРЕДСТАВЛЕНИЕ(Рег.Source) КАК Source,
    Рег.Расхождение КАК Расхождение,
    СУММА(Рег.СуммаНачальныйОстаток) КАК НО,
    СУММА(Рег.СуммаКонечныйОстаток) КАК КО,
    КОЛИЧЕСТВО(*) КАК Строк
ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Рег
ГДЕ Рег.Статья = &ОТ
СГРУППИРОВАТЬ ПО
    Рег.Регистратор.Месяц, Рег.Регистратор.Номер, Рег.Регистратор.Проведен,
    ПРЕДСТАВЛЕНИЕ(Рег.Source), Рег.Расхождение
УПОРЯДОЧИТЬ ПО Месяц
"""
ta = qa.Выполнить().Выгрузить()
reg_by_month = {}
for r in ta:
    m = r.Месяц.strftime("%Y-%m") if r.Месяц else "?"
    reg_by_month[m] = reg_by_month.get(m, 0.0) + f(r.КО)
    print(f"  {m} док №{str(r.Номер).strip()} Провед={r.Проведен} | Source='{r.Source}' Расхожд={r.Расхождение} "
          f"| строк={int(f(r.Строк)):>4} | НО={f(r.НО):>16,.2f} | КО={f(r.КО):>16,.2f}")
print("  --- ИТОГО КО по месяцам (регистр) ---")
for m in sorted(reg_by_month):
    print(f"    {m}: {reg_by_month[m]:>16,.2f}")

# === B. ПАП источник истины — Остаток на конец каждого месяца ===
print("\n=== B. РегистрНакопления.ПрочиеАктивыПассивы (Статья=ОТ, Источник=ПустаяСсылка) — Остаток ===")
pap_by_month = {}
for label, dt in month_ends:
    qb = erp.NewObject("Запрос")
    qb.УстановитьПараметр("ОТ", ОТ)
    qb.УстановитьПараметр("Дата", dt)
    qb.Текст = """
    ВЫБРАТЬ
        СУММА(Ост.СуммаОстаток) КАК Остаток
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы.Остатки(
        &Дата,
        Статья = &ОТ
        И Источник = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.ПустаяСсылка)) КАК Ост
    """
    rb = qb.Выполнить().Выгрузить()
    ost = f(rb[0].Остаток) if rb.Количество() > 0 else 0.0
    pap_by_month[label] = ost
    print(f"  {label} (на {dt:%d.%m.%Y}): Остаток ПАП = {ost:>16,.2f}")

# === C. SQL Fact_Balance (OlapBASERP) ===
print("\n=== C. OlapBASERP.Fact_Balance (Статья=Оплата труда) ===")
sql_by_month = {}
try:
    import pyodbc
    cn = pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;DATABASE=OlapBASERP;"
        "UID=sa;PWD=Brw739182465!", timeout=10)
    cur = cn.cursor()
    # схема колонок Fact_Balance
    cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Fact_Balance' ORDER BY ORDINAL_POSITION")
    cols = [row[0] for row in cur.fetchall()]
    print(f"  Fact_Balance колонки: {cols}")
    cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Dim_PAP_Articles' ORDER BY ORDINAL_POSITION")
    print(f"  Dim_PAP_Articles колонки: {[row[0] for row in cur.fetchall()]}")
    cur.execute("""
        SELECT d.PAP_Article_Name, CONVERT(varchar(7), f.Period, 120) AS Mon,
               f.Source, SUM(f.Sum_Close) AS Close, COUNT(*) AS rows_
        FROM Fact_Balance f
        JOIN Dim_PAP_Articles d ON f.PAP_Article_ID = d.PAP_Article_ID
        WHERE d.PAP_Article_Name = N'Оплата труда'
        GROUP BY d.PAP_Article_Name, CONVERT(varchar(7), f.Period, 120), f.Source
        ORDER BY Mon, f.Source
    """)
    for row in cur.fetchall():
        m = row.Mon
        sql_by_month[m] = sql_by_month.get(m, 0.0) + f(row.Close)
        print(f"  {m} | Source='{row.Source}' | строк={int(row.rows_):>4} | Close={f(row.Close):>16,.2f}")
    print("  --- ИТОГО Close по месяцам (Fact_Balance) ---")
    for m in sorted(sql_by_month):
        print(f"    {m}: {sql_by_month[m]:>16,.2f}")
    # последняя загрузка ETL
    cur.execute("SELECT MAX(Loaded_At) FROM Fact_Balance")
    print(f"  Fact_Balance MAX(Loaded_At) = {cur.fetchone()[0]}")
except Exception as e:
    print(f"  [pyodbc FAIL] {e}")

# === Сводка расхождений ===
print("\n=== СВОДКА: КО регистр vs ПАП vs Fact_Balance(PowerBI) ===")
print(f"  {'Месяц':<8} | {'Регистр КО':>16} | {'ПАП Остаток':>16} | {'Fact_Balance':>16} | {'Δ(PBI-Рег)':>16}")
allm = sorted(set(list(reg_by_month) + list(pap_by_month) + list(sql_by_month)))
for m in allm:
    reg = reg_by_month.get(m, 0.0)
    pap = pap_by_month.get(m, 0.0)
    pbi = sql_by_month.get(m, 0.0)
    print(f"  {m:<8} | {reg:>16,.2f} | {pap:>16,.2f} | {pbi:>16,.2f} | {pbi-reg:>16,.2f}")
