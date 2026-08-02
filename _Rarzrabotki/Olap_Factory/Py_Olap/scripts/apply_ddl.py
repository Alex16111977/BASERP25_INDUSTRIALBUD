# -*- coding: utf-8 -*-
"""
Применяет scripts/ddl_olap_factory.sql к базе OlapFactory + сеет справочные значения,
которые нельзя держать в .sql (кириллица: sqlcmd портит кодировку — вставляем pyodbc-параметрами).

  * DDL батчами по "GO", autocommit;
  * Calendar — ЗЕРКАЛО живой OlapBASERP.dbo.Calendar (43 колонки, 2024-01-01 .. 2027-12-31):
    структура читается из INFORMATION_SCHEMA соседней базы и воспроизводится 1:1, данные
    копируются межбазовым INSERT ... SELECT. Так витрина совпадает с эталоном, который видно
    в PowerBi/Industrial/PL.pbix, и не устаревает при доработке эталона.
    🔴 Почему не в .sql: у эталона есть колонка с кириллическим именем «Квартал», а
    ddl_olap_factory.sql держится ASCII-only. 🔴 Почему не перезапуском
    Olap/Ai_Olap/scripts/calendar_dim_olapbaserp.sql: он устарел на 3 колонки и содержит
    USE OlapBASERP — писать в чужую базу нельзя (читать можно).
  * Dim_RowTypes — порядок ОБЯЗАН совпадать с _EnumOrder перечисления
    А_ТипыСтрокПланФактПроизводство (и с FROZEN_ENUMS в ai_olap/transformers/enum_resolver.py).

Запуск: .venv/Scripts/python.exe scripts/apply_ddl.py
"""
import os, sys, datetime as dt
from pathlib import Path

import pyodbc
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding='utf-8')
ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / ".env")
DSN = os.environ["OLAP_SQL_DSN"]

ЭТАЛОН_БД = "OlapBASERP"          # соседний контур PnL/Cashflow/Balance — ТОЛЬКО ЧТЕНИЕ
ЭТАЛОН_ТАБЛИЦА = "Calendar"

МЕСЯЦЫ = ["січень", "лютий", "березень", "квітень", "травень", "червень",
          "липень", "серпень", "вересень", "жовтень", "листопад", "грудень"]
МЕСЯЦЫ_КОР = ["січ", "лют", "бер", "кві", "тра", "чер", "лип", "сер", "вер", "жов", "лис", "гру"]
ДНИ = ["понеділок", "вівторок", "середа", "четвер", "п'ятниця", "субота", "неділя"]

# порядок = _EnumOrder в 1С; RowType_Group делит «читать за период» / «читать целиком»
ТИПЫ_СТРОК = [
    ("План",       0, "Незалежний"),
    ("Факт",       1, "Період"),
    ("Виконання",  2, "Період"),
    ("ПланНаФакт", 3, "Незалежний"),
]

# запасной 16-колоночный календарь: используется ТОЛЬКО если эталонная база недоступна
ЛЕГАСИ_DDL = """
CREATE TABLE Calendar (
    date_             datetime NOT NULL PRIMARY KEY,
    year_             int NOT NULL,
    quarter_          int NOT NULL,
    month_            int NOT NULL,
    day_              int NOT NULL,
    week_             int NOT NULL,
    week_day          int NOT NULL,
    month_name_ua     nvarchar(40) NULL,
    month_short_ua    nvarchar(10) NULL,
    week_day_ua       nvarchar(40) NULL,
    year_month        varchar(7) NOT NULL,
    Period_Month      date NOT NULL,
    Period_Quarter    date NOT NULL,
    Period_Year       date NOT NULL,
    Year_Month_Sort   int NOT NULL,
    weekend           bit NOT NULL
)"""


def применить_ddl(cur):
    sql = (ROOT / "scripts" / "ddl_olap_factory.sql").read_text(encoding="utf-8")
    батчи = [b.strip() for b in sql.split("\nGO") if b.strip()]
    for i, b in enumerate(батчи, 1):
        try:
            cur.execute(b)
        except Exception as exc:
            print(f"  батч {i}/{len(батчи)} FAIL: {exc}")
            print(b[:300])
            raise
    print(f"[OK] DDL применён: {len(батчи)} батчей")


# ------------------------------------------------------------------ Calendar
def _тип(тип, длина, точность, масштаб):
    if тип in ("nvarchar", "varchar", "char", "nchar", "binary", "varbinary"):
        return f"{тип}(max)" if длина == -1 else f"{тип}({длина})"
    if тип in ("decimal", "numeric"):
        return f"{тип}({точность},{масштаб})"
    return тип


def структура(cur, база, таблица):
    """[(имя, тип_с_длиной, признак_NULL)] в порядке колонок; [] если таблицы нет."""
    cur.execute(f"""
        SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH,
               NUMERIC_PRECISION, NUMERIC_SCALE, IS_NULLABLE
        FROM {база}.INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = ?
        ORDER BY ORDINAL_POSITION""", таблица)
    return [(r[0], _тип(r[1], r[2], r[3], r[4]), r[5] == "YES") for r in cur.fetchall()]


def база_доступна(cur, база):
    cur.execute("SELECT COUNT(*) FROM sys.databases WHERE name = ?", база)
    return bool(cur.fetchone()[0])


def синхронизировать_календарь(cur):
    """Зеркалит OlapBASERP.dbo.Calendar в OlapFactory: структура 1:1 + все строки."""
    if not база_доступна(cur, ЭТАЛОН_БД):
        return _легаси_календарь(cur, f"база {ЭТАЛОН_БД} не найдена на сервере")
    эталон = структура(cur, ЭТАЛОН_БД, ЭТАЛОН_ТАБЛИЦА)
    if not эталон:
        return _легаси_календарь(cur, f"нет таблицы {ЭТАЛОН_БД}.dbo.{ЭТАЛОН_ТАБЛИЦА}")

    наша = структура(cur, "OlapFactory", "Calendar")
    if наша != эталон:
        if наша:
            print(f"  структура Calendar отличается от эталона "
                  f"({len(наша)} колонок против {len(эталон)}) — пересоздаю")
            cur.execute("DROP TABLE dbo.Calendar")
        колонки = ",\n    ".join(
            f"[{и}] {т} {'NULL' if н else 'NOT NULL'}" for и, т, н in эталон)
        cur.execute(f"CREATE TABLE dbo.Calendar (\n    {колонки},\n"
                    f"    CONSTRAINT PK_Calendar PRIMARY KEY ([date_])\n)")
        print(f"[OK] Calendar создан по эталону {ЭТАЛОН_БД}: {len(эталон)} колонок")

    список = ", ".join(f"[{и}]" for и, _, _ in эталон)
    cur.execute("DELETE FROM dbo.Calendar")
    cur.execute(f"INSERT INTO dbo.Calendar ({список}) "
                f"SELECT {список} FROM {ЭТАЛОН_БД}.dbo.{ЭТАЛОН_ТАБЛИЦА}")
    cur.execute("SELECT COUNT(*), MIN(date_), MAX(date_) FROM dbo.Calendar")
    n, dmin, dmax = cur.fetchone()
    print(f"[OK] Calendar скопирован: {n} дней ({dmin:%Y-%m-%d} .. {dmax:%Y-%m-%d})")


def _легаси_календарь(cur, причина):
    """Аварийный путь: эталона нет — сеем прежний 16-колоночный календарь.
    Модель Power BI при этом с PL.pbix НЕ совпадёт: в ней 43 колонки эталона."""
    print(f"  !! ЭТАЛОН НЕДОСТУПЕН ({причина}) — откат на прежнюю 16-колоночную сеялку.")
    print(f"  !! Модель PBIP рассчитана на 43 колонки эталона и работать не будет,")
    print(f"  !! пока Calendar не скопируют из {ЭТАЛОН_БД}. Это НЕ штатный режим.")
    if not структура(cur, "OlapFactory", "Calendar"):
        cur.execute(ЛЕГАСИ_DDL)
    cur.execute("SELECT COUNT(*) FROM Calendar")
    if cur.fetchone()[0]:
        print("  Calendar уже заполнен — пропуск")
        return
    строки = []
    d = dt.date(2024, 1, 1)
    конец = dt.date(2028, 12, 31)
    while d <= конец:
        нед = d.isoweekday()
        кв = (d.month - 1) // 3 + 1
        строки.append((
            dt.datetime(d.year, d.month, d.day), d.year, кв, d.month, d.day,
            d.isocalendar()[1], нед,
            МЕСЯЦЫ[d.month - 1], МЕСЯЦЫ_КОР[d.month - 1], ДНИ[нед - 1],
            f"{d.year}-{d.month:02d}",
            dt.date(d.year, d.month, 1), dt.date(d.year, (кв - 1) * 3 + 1, 1),
            dt.date(d.year, 1, 1), d.year * 100 + d.month, 1 if нед >= 6 else 0,
        ))
        d += dt.timedelta(days=1)
    cur.fast_executemany = True
    cur.executemany(
        "INSERT INTO Calendar (date_, year_, quarter_, month_, day_, week_, week_day, "
        "month_name_ua, month_short_ua, week_day_ua, year_month, Period_Month, Period_Quarter, "
        "Period_Year, Year_Month_Sort, weekend) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", строки)
    print(f"[OK] Calendar засеян ЛЕГАСИ: {len(строки)} дней (2024-01-01 .. 2028-12-31)")


def засеять_типы(cur):
    cur.execute("SELECT COUNT(*) FROM Dim_RowTypes")
    if cur.fetchone()[0]:
        print("  Dim_RowTypes уже заполнен — пропуск")
        return
    cur.executemany("INSERT INTO Dim_RowTypes (RowType, RowType_Sort, RowType_Group) VALUES (?,?,?)",
                    ТИПЫ_СТРОК)
    print(f"[OK] Dim_RowTypes засеян: {len(ТИПЫ_СТРОК)} значений")


def main():
    cn = pyodbc.connect(DSN, autocommit=True)
    cur = cn.cursor()
    применить_ddl(cur)
    синхронизировать_календарь(cur)
    засеять_типы(cur)
    cur.execute("""SELECT t.name, ISNULL(p.rows,0) FROM sys.tables t
                   LEFT JOIN sys.partitions p ON p.object_id=t.object_id AND p.index_id IN (0,1)
                   ORDER BY t.name""")
    print("\nТаблицы OlapFactory:")
    for имя, n in cur.fetchall():
        print(f"   {имя:32} {n:>8} строк")
    cn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
