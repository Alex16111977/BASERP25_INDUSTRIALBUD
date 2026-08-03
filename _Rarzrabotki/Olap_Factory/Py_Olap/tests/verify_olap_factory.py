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


# Паритет разрезов с эталонным отчётом А_ПланФактныйПроизводствоПолный.
# Ожидание — НЕ «заполнено везде», а «заполнено ровно там, где заполнено в отчёте»:
#   ПодразделениеИсполнитель — факт (табель) и виконання; у плана и ETC пусто;
#   ПодразделениеОбъекта     — план и виконання (реквизит карточки СС); у факта пусто
#                              («котёл» на итоге подразделения), у ETC пусто by design.
# Пустая колонка там, где ожидается заполнение, означает разорванную цепочку
# Документ → карточка СС и молча пустой разрез в Power BI.
ОЖИДАНИЕ_РАЗРЕЗОВ = {
    #  тип строки   : (есть ли виконавець, есть ли обʼєкт)
    "План":          (False, True),
    "Факт":          (True,  False),
    "Виконання":     (True,  True),
    "ПланНаФакт":    (False, False),
}


def проверить_разрезы_подразделений(cur):
    cur.execute("""
        SELECT RowType, COUNT(*),
               SUM(CASE WHEN ExecutorDept_ID IS NOT NULL THEN 1 ELSE 0 END),
               SUM(CASE WHEN ObjectDept_ID   IS NOT NULL THEN 1 ELSE 0 END)
        FROM vw_Fact_PlanFact GROUP BY RowType ORDER BY RowType""")
    print(f"   {'Тип строки':14} {'Строк':>6} {'виконавець':>11} {'обʼєкт':>8}   ожидание")
    for тип, всего, исп, обj in cur.fetchall():
        ждём = ОЖИДАНИЕ_РАЗРЕЗОВ.get(тип)
        if ждём is None:
            провалы.append(f"неизвестный ТипСтроки {тип}")
            continue
        факт = (исп > 0, обj > 0)
        ок = факт == ждём
        print(f"   {тип:14} {всего:>6} {исп:>11} {обj:>8}   "
              f"{'OK' if ок else f'ЖДАЛИ {ждём}, ПОЛУЧИЛИ {факт}'}")
        if not ок:
            провалы.append(f"разрезы {тип}: ждали {ждём}, получили {факт}")

    for кол_имя in ("ExecutorDept_ID", "ObjectDept_ID"):
        cur.execute(f"""SELECT COUNT(*) FROM vw_Fact_PlanFact v
                        LEFT JOIN Dim_Departments d ON d.Department_ID = v.{кол_имя}
                        WHERE v.{кол_имя} IS NOT NULL AND d.Department_ID IS NULL""")
        n = cur.fetchone()[0]
        print(f"   сироты {кол_имя:16} -> Dim_Departments  {n}")
        if n:
            провалы.append(f"FK-сироты {кол_имя}: {n}")

    # объект обязан дробить план ровно внутри подразделения, а не поверх него:
    # сумма плана по объектам одного подразделения = плану этого подразделения
    cur.execute("""
        SELECT d.Department_Name, SUM(v.PlanUAH) AS ПланПодр,
               SUM(CASE WHEN v.ObjectDept_ID IS NOT NULL THEN v.PlanUAH ELSE 0 END) AS ПланСОбъектом
        FROM vw_Fact_PlanFact v
        LEFT JOIN Dim_Departments d ON d.Department_ID = v.Department_ID
        WHERE v.RowType = 'План'
        GROUP BY d.Department_Name ORDER BY d.Department_Name""")
    print("   план по подразделению == план, разложенный по объектам:")
    for имя, всего, с_объектом in cur.fetchall():
        d = abs(float(всего) - float(с_объектом))
        print(f"      {имя:28} {float(всего):>15,.2f}   {'OK' if d < 0.005 else f'Δ={d:,.2f}'}")
        if d >= 0.005:
            провалы.append(f"план {имя}: без объекта осталось {d:,.2f}")

    # 🔴 Число объектов НЕ является глобальным инвариантом витрины: два объекта (15м/30м) —
    # свойство ровно одного проекта. Из 5 активных карточек СС две принадлежат «МД IRS 2026»,
    # у остальных трёх подразделений по одной карточке, и объект там совпадает с подразделением.
    # Регистр — объединение по-подразделенческих снимков, поэтому витрина законно может содержать
    # ровно один объект. Требовать «объектов >= 2» значит краснеть на исправных данных.
    # Осмысленная проверка — поштучная: у подразделения объектов не больше, чем карточек СС,
    # и хотя бы один. Вырождение join-а (всем строкам достался один и тот же объект) ловится
    # именно этим сравнением, а не глобальным порогом.
    cur.execute("""
        SELECT d.Department_Name,
               COUNT(DISTINCT f.Document_ID)   AS Карточек,
               COUNT(DISTINCT f.ObjectDept_ID) AS Обьектов
        FROM vw_Fact_PlanFact f
        JOIN Dim_Documents dd ON dd.Document_ID = f.Document_ID
                             AND dd.Document_Type = N'Структура себестоимости'
        LEFT JOIN Dim_Departments d ON d.Department_ID = f.Department_ID
        WHERE f.RowType = 'План'
        GROUP BY d.Department_Name ORDER BY d.Department_Name""")
    print("   объекты и карточки СС по подразделениям:")
    for имя, карточек, объектов in cur.fetchall():
        пометка = "OK"
        if объектов < 1:
            пометка = "НЕТ ОБЪЕКТА"
            провалы.append(f"{имя}: план без объекта при {карточек} карточках СС")
        elif объектов > карточек:
            пометка = "БОЛЬШЕ, ЧЕМ КАРТОЧЕК"
            провалы.append(f"{имя}: объектов {объектов} при {карточек} карточках СС")
        elif карточек > 1 and объектов == 1:
            # законно, если две карточки указывают на один объект — но стоит увидеть глазами
            пометка = "внимание: карточек несколько, объект один"
        print(f"      {str(имя):28} карточек {карточек}, объектов {объектов}   {пометка}")

    cur.execute("""SELECT COUNT(DISTINCT ObjectDept_ID) FROM vw_Fact_PlanFact
                   WHERE ObjectDept_ID IS NOT NULL""")
    n_об = cur.fetchone()[0]
    cur.execute("""SELECT COUNT(DISTINCT ExecutorDept_ID) FROM vw_Fact_PlanFact
                   WHERE ExecutorDept_ID IS NOT NULL""")
    n_исп = cur.fetchone()[0]
    print(f"   различных: обʼєктів {n_об}, виконавців {n_исп}")


def main():
    cn = pyodbc.connect(OLAP)
    cur = cn.cursor()
    erp = win32com.client.Dispatch("V83.COMConnector").Connect(COM)

    print("=" * 92)
    print("6.1 — ETL_Runs")
    print("=" * 92)
    # 🔴 Проваливает приёмку только ПОСЛЕДНИЙ прогон каждого скрипта. Упавший, но затем
    # исправленный прогон остаётся в журнале как исторический факт — удалять его, чтобы
    # тест позеленел, значит подделать журнал; блокировать им приёмку — значит держать
    # красный тест до тех пор, пока запись не вытеснится из TOP 5.
    cur.execute("""SELECT r.Run_ID, r.Script, r.Status, r.Rows_Loaded, r.Error,
                          CASE WHEN r.Run_ID = (SELECT MAX(Run_ID) FROM ETL_Runs x
                                                WHERE x.Script = r.Script) THEN 1 ELSE 0 END
                   FROM ETL_Runs r ORDER BY r.Run_ID DESC
                   OFFSET 0 ROWS FETCH NEXT 5 ROWS ONLY""")
    for r in cur.fetchall():
        метка = "последний" if r[5] else "исторический"
        print(f"   run {r[0]:>3}  {r[1]:30} {r[2]:8} rows={r[3]}  ({метка})")
        if r[2] != "Success" and r[5]:
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

    print()
    print("=" * 92)
    print("6.9 — разрезы подразделений: виконавець и обʼєкт")
    print("=" * 92)
    проверить_разрезы_подразделений(cur)

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
