# -*- coding: utf-8 -*-
"""
Приёмка витрины OlapFactory (этап 6 плана).

  6.1 ETL_Runs — последние прогоны Success;
  6.2 число строк Fact == число записей РегистрСведений.А_ПланФактПроизводство_Свод;
  6.3 суммы всех 8 ресурсов SQL == 1С до копейки (COM-запрос к регистру);
  6.4 разрез по подразделениям SQL == 1С;
  6.5 FK-сироты по всем Dim == 0;
  6.6 RowType: NULL нет, состав совпадает с Dim_RowTypes;
  6.7 Period в человеческом диапазоне (не 4026) и покрыт Calendar.

Запуск: .venv/Scripts/python.exe tests/verify_olap_factory.py
"""
import os, sys
from pathlib import Path

import pyodbc, win32com.client
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding='utf-8')
ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / ".env")
OLAP = os.environ["OLAP_SQL_DSN"]
COM = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

РЕСУРСЫ = [
    ("PlanHours",   "ПланЧасы"),
    ("FactHours",   "ФактЧасы"),
    ("EarnedHours", "ВиконанняЧасы"),
    ("PlanQty",     "ПланКол"),
    ("FactQty",     "ФактКол"),
    ("PlanUAH",     "ПланГрн"),
    ("FactUAH",     "ФактГрн"),
    ("ETC_UAH",     "ПланНаФактГрн"),
]
провалы = []
ЭТАЛОН_БД = "OlapBASERP"   # соседний контур: только читаем


def _структура(cur, база):
    cur.execute(f"""
        SELECT COLUMN_NAME, DATA_TYPE, ISNULL(CHARACTER_MAXIMUM_LENGTH, -9), IS_NULLABLE
        FROM {база}.INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'Calendar'
        ORDER BY ORDINAL_POSITION""")
    return [tuple(r) for r in cur.fetchall()]


def проверить_календарь(cur):
    """Состав и типы колонок сверяются ПРОГРАММНО, а не глазами: у витрины Calendar обязан
    быть точной копией OlapBASERP.dbo.Calendar (43 колонки), иначе модель Power BI разойдётся
    с эталоном PL.pbix — и разойдётся молча, без ошибки при открытии."""
    эталон, наша = _структура(cur, ЭТАЛОН_БД), _структура(cur, "OlapFactory")
    print(f"   колонок: витрина {len(наша)}   эталон {len(эталон)}")
    if not эталон:
        провалы.append(f"нет {ЭТАЛОН_БД}.dbo.Calendar — сверять не с чем")
        return
    лишние = [c[0] for c in наша if c[0] not in {e[0] for e in эталон}]
    нехватка = [e[0] for e in эталон if e[0] not in {c[0] for c in наша}]
    типы = [f"{e[0]}: витрина {c[1]}({c[2]}) {c[3]} vs эталон {e[1]}({e[2]}) {e[3]}"
            for e, c in zip(эталон, наша) if e[0] == c[0] and e[1:] != c[1:]]
    for имя, список in (("лишние", лишние), ("отсутствуют", нехватка), ("типы", типы)):
        print(f"   {имя}: {список or 'нет'}")
        for x in список:
            провалы.append(f"Calendar {имя}: {x}")
    if наша != эталон and not (лишние or нехватка or типы):
        провалы.append("Calendar: порядок колонок отличается от эталона")

    cur.execute(f"""SELECT (SELECT COUNT(*) FROM OlapFactory.dbo.Calendar),
                           (SELECT COUNT(*) FROM {ЭТАЛОН_БД}.dbo.Calendar),
                           (SELECT MIN(date_) FROM OlapFactory.dbo.Calendar),
                           (SELECT MAX(date_) FROM OlapFactory.dbo.Calendar)""")
    n, nэ, dmin, dmax = cur.fetchone()
    print(f"   строк: витрина {n}   эталон {nэ}   период {dmin:%Y-%m-%d} .. {dmax:%Y-%m-%d}")
    if n != nэ:
        провалы.append(f"Calendar строк {n} vs эталон {nэ}")

    # значения сходятся до символа: сравниваем по всем колонкам, а не только по date_
    список = ", ".join(f"[{c[0]}]" for c in эталон)
    cur.execute(f"""SELECT COUNT(*) FROM (
                        SELECT {список} FROM OlapFactory.dbo.Calendar
                        EXCEPT
                        SELECT {список} FROM {ЭТАЛОН_БД}.dbo.Calendar) x""")
    d = cur.fetchone()[0]
    print(f"   строк, отличающихся от эталона хотя бы одним значением: {d}")
    if d:
        провалы.append(f"Calendar: {d} строк не совпадают с эталоном")


def main():
    cn = pyodbc.connect(OLAP)
    cur = cn.cursor()
    erp = win32com.client.Dispatch("V83.COMConnector").Connect(COM)

    print("=" * 92)
    print("6.1 — ETL_Runs")
    print("=" * 92)
    cur.execute("SELECT TOP 5 Run_ID, Script, Status, Rows_Loaded, Error FROM ETL_Runs ORDER BY Run_ID DESC")
    for r in cur.fetchall():
        print(f"   run {r[0]:>3}  {r[1]:30} {r[2]:8} rows={r[3]}")
        if r[2] != "Success":
            провалы.append(f"ETL_Runs {r[0]} {r[1]}: {r[2]} / {r[4]}")

    print()
    print("=" * 92)
    print("6.2-6.3 — строки и суммы: SQL vs 1С")
    print("=" * 92)
    q = erp.NewObject("Запрос")
    q.Текст = ("ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К, " +
               ", ".join(f"СУММА(Р.{ру}) КАК {ан}" for ан, ру in РЕСУРСЫ) +
               " ИЗ РегистрСведений.А_ПланФактПроизводство_Свод КАК Р")
    од = q.Execute().Выгрузить().Получить(0)
    n1c = int(од.К)

    cur.execute("SELECT COUNT(*), " + ", ".join(f"SUM({a})" for a, _ in РЕСУРСЫ) +
                " FROM Fact_PlanFactProizvodstvo")
    строка = cur.fetchone()
    nsql = int(строка[0])
    print(f"   строк: SQL {nsql}   1С {n1c}   {'OK' if nsql == n1c else 'РАСХОЖДЕНИЕ'}")
    if nsql != n1c:
        провалы.append(f"строк SQL {nsql} vs 1С {n1c}")
    for i, (ан, ру) in enumerate(РЕСУРСЫ, start=1):
        vsql = float(строка[i] or 0)
        v1c = float(getattr(од, ан) or 0)
        d = abs(vsql - v1c)
        print(f"   {ан:12} SQL {vsql:>18,.3f}   1С {v1c:>18,.3f}   {'OK' if d < 0.005 else f'Δ={d:,.3f}'}")
        if d >= 0.005:
            провалы.append(f"{ан}: SQL {vsql} vs 1С {v1c}")

    print()
    print("=" * 92)
    print("6.4 — разрез по подразделениям")
    print("=" * 92)
    q = erp.NewObject("Запрос")
    q.Текст = ("ВЫБРАТЬ Р.Подразделение.Наименование КАК Подр, СУММА(Р.ПланГрн) КАК ПланГрн, "
               "СУММА(Р.ФактГрн) КАК ФактГрн, СУММА(Р.ПланНаФактГрн) КАК ETC "
               "ИЗ РегистрСведений.А_ПланФактПроизводство_Свод КАК Р "
               "СГРУППИРОВАТЬ ПО Р.Подразделение.Наименование")
    т = q.Execute().Выгрузить()
    из1с = {str(т.Получить(i).Подр): (float(т.Получить(i).ПланГрн), float(т.Получить(i).ФактГрн),
                                      float(т.Получить(i).ETC)) for i in range(т.Количество())}
    cur.execute("""SELECT d.Department_Name, SUM(f.PlanUAH), SUM(f.FactUAH), SUM(f.ETC_UAH)
                   FROM Fact_PlanFactProizvodstvo f
                   LEFT JOIN Dim_Departments d ON d.Department_ID = f.Department_ID
                   GROUP BY d.Department_Name ORDER BY d.Department_Name""")
    for имя, p, fa, e in cur.fetchall():
        эт = из1с.get(имя)
        if эт is None:
            print(f"   {имя:26} НЕТ В 1С")
            провалы.append(f"подразделение {имя} отсутствует в 1С")
            continue
        d = max(abs(float(p) - эт[0]), abs(float(fa) - эт[1]), abs(float(e) - эт[2]))
        print(f"   {имя:26} План {float(p):>15,.2f}  Факт {float(fa):>15,.2f}  ETC {float(e):>14,.2f}   "
              f"{'OK' if d < 0.005 else f'Δ={d:,.2f}'}")
        if d >= 0.005:
            провалы.append(f"подразделение {имя}: Δ={d}")

    print()
    print("=" * 92)
    print("6.5-6.7 — целостность")
    print("=" * 92)
    fk = [("Department_ID", "Dim_Departments", "Department_ID"),
          ("ExecutorDept_ID", "Dim_Departments", "Department_ID"),
          ("Organization_ID", "Dim_Organizations", "Organization_ID"),
          ("Stage_ID", "Dim_Stages", "Stage_ID"),
          ("Work_ID", "Dim_Items", "Item_ID"),
          ("CommonName_ID", "Dim_CommonNames", "CommonName_ID"),
          ("Unit_ID", "Dim_Units", "Unit_ID"),
          ("RowType", "Dim_RowTypes", "RowType")]
    for кол, таб, ключ in fk:
        cur.execute(f"SELECT COUNT(*) FROM Fact_PlanFactProizvodstvo f LEFT JOIN {таб} d "
                    f"ON d.{ключ} = f.{кол} WHERE f.{кол} IS NOT NULL AND d.{ключ} IS NULL")
        n = cur.fetchone()[0]
        print(f"   сироты {кол:16} -> {таб:20} {n}")
        if n:
            провалы.append(f"FK-сироты {кол}->{таб}: {n}")

    cur.execute("SELECT COUNT(*) FROM Fact_PlanFactProizvodstvo WHERE RowType IS NULL OR RowType = ''")
    n = cur.fetchone()[0]
    print(f"   RowType пустых: {n}")
    if n:
        провалы.append(f"RowType пустых: {n}")
    cur.execute("SELECT RowType, COUNT(*) FROM Fact_PlanFactProizvodstvo GROUP BY RowType ORDER BY RowType")
    print("   состав RowType:", {r[0]: r[1] for r in cur.fetchall()})

    cur.execute("SELECT MIN(Period), MAX(Period), MIN(Period_Month), MAX(Period_Month) FROM Fact_PlanFactProizvodstvo")
    mn, mx, pm1, pm2 = cur.fetchone()
    print(f"   Period: {mn} .. {mx}   Period_Month: {pm1} .. {pm2}")
    if mn.year > 2100 or mx.year > 2100:
        провалы.append(f"офсет +2000 не снят: {mn}..{mx}")

    cur.execute("""SELECT COUNT(*) FROM Fact_PlanFactProizvodstvo f
                   LEFT JOIN Calendar c ON c.date_ = f.Period WHERE c.date_ IS NULL""")
    n = cur.fetchone()[0]
    print(f"   строк без связи с Calendar: {n}")
    if n:
        провалы.append(f"вне Calendar: {n}")

    print()
    print("=" * 92)
    print("6.8 — Calendar == эталон OlapBASERP (тот, что виден в PL.pbix)")
    print("=" * 92)
    проверить_календарь(cur)

    cn.close()
    if провалы:
        print("\nFAIL:")
        for p in провалы:
            print("   ", p)
        return 1
    print("\nПРИЁМКА ВИТРИНЫ OlapFactory ПРОЙДЕНА.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
