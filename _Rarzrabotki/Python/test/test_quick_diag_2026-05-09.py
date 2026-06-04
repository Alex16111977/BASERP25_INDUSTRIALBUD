# -*- coding: utf-8 -*-
"""Швидка діагностика — подивитися структуру і стан."""
import sys
import win32com.client
import pyodbc

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

print("=" * 100)
print("Бюджет 000002597 — поточний стан")
print("=" * 100)

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Д.Ссылка КАК Ссылка,
    Д.Дата КАК Дата,
    Д.Проведен КАК Проведен,
    Д.ПометкаУдаления КАК Удалено,
    Д.ВидВерсии КАК Версия,
    Д.ДокументОснование КАК Основание
ИЗ Документ.А_БюджетМесяц КАК Д
ГДЕ Д.Номер = "000002597"
"""
res = q.Execute().Выгрузить()
for i in range(res.Количество()):
    r = res.Получить(i)
    print(f"Дата: {S(r.Дата)}")
    print(f"Проведен: {r.Проведен}")
    print(f"Удалено: {r.Удалено}")
    print(f"Версия: {S(r.Версия)}")
    print(f"Основание: {S(r.Основание)}")

print()
print("=" * 100)
print("Записи у А_БюджетыНаМесяц по Бюджет 000002597 — ВСЕ записи")
print("=" * 100)

q2 = erp.NewObject("Запрос")
q2.Text = """
ВЫБРАТЬ ПЕРВЫЕ 10
    Р.Период, Р.Месяц, Р.Подразделение, Р.СтатьяДвиженияДенежныхСредств,
    Р.ВидПериода, Р.Сумма, Р.Активность
ИЗ РегистрНакопления.А_БюджетыНаМесяц КАК Р
ГДЕ Р.Регистратор.Номер = "000002597"
    И Р.Сумма < 0
УПОРЯДОЧИТЬ ПО Р.НомерСтроки
"""
res2 = q2.Execute().Выгрузить()
print(f"Знайдено сторно-рядків (Сумма<0): {res2.Количество()}")
for i in range(min(5, res2.Количество())):
    r = res2.Получить(i)
    print(f"  {S(r.Период)} | Месяц={r.Месяц.strftime('%d.%m.%Y')} | Подр={S(r.Подразделение)[:30]} | Стат={S(r.СтатьяДвиженияДенежныхСредств)[:30]} | ВидПер={S(r.ВидПериода)} | Сумма={r.Сумма} | Активность={r.Активность}")


print()
print("=" * 100)
print("Поля регістра А_ОтчетDDS_Свод (як визначено у конфігурації)")
print("=" * 100)

# Get the FieldList for А_ОтчетDDS_Свод
mng = erp.РегистрыСведений.А_ОтчетDDS_Свод
nq = erp.NewObject("Запрос")
nq.Text = """
ВЫБРАТЬ ПЕРВЫЕ 1 *
ИЗ РегистрСведений.А_ОтчетDDS_Свод
"""
nres = nq.Execute().Выгрузить()
print(f"Колонок: {nres.Колонки.Количество()}")
for i in range(nres.Колонки.Количество()):
    col = nres.Колонки.Получить(i)
    print(f"  {col.Имя}")


print()
print("=" * 100)
print("Колонки SQL Fact_Cashflow")
print("=" * 100)

try:
    sql = pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=localhost;DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;"
        "TrustServerCertificate=yes;"
    )
    cur = sql.cursor()
    cur.execute("""
        SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME='Fact_Cashflow' ORDER BY ORDINAL_POSITION
    """)
    cols = [r[0] for r in cur.fetchall()]
    for c in cols:
        print(f"  {c}")
    sql.close()
except Exception as e:
    print(f"FAIL: {e}")
