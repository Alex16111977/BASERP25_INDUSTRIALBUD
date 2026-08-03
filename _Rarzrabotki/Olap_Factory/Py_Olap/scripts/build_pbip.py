# -*- coding: utf-8 -*-
"""
Генерирует Power BI проект (PBIP: семантическая модель TMDL + пустой отчёт) поверх витрины
OlapFactory. Формат тот же, что у существующего _Rarzrabotki/PowerBi/Industrial/PL.pbip —
Power BI Desktop 2.153 открывает его напрямую и позволяет сохранить как .pbix.

Почему PBIP, а не .pbix: .pbix — бинарный контейнер с моделью в формате ABF, его нельзя
собрать скриптом. PBIP — текст (TMDL + JSON), полностью воспроизводим и версионируется в git.

Каждая таблица тянется NATIVE QUERY с ЯВНЫМ списком колонок: коннектор Sql.Database кеширует
схему, и при обычном [Schema=,Item=] новая колонка после ALTER не появляется даже после Full
Refresh (грабля проекта pbix_sql_database_schema_cache).

🔴 План и «План на факт» в мерах обёрнуты в REMOVEFILTERS('Calendar') — это переносит в DAX
правило контура: факт и виконання считаются за выбранный период, план и ETC видны целиком.

Запуск: .venv/Scripts/python.exe scripts/build_pbip.py
"""
import json, re, sys, uuid, shutil
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

ВЫХОД = Path(r"C:/Configuration_downloads/BASERP25/_Rarzrabotki/Olap_Factory")
ИМЯ = "ПланФактВиробництва"
СЕРВЕР, БАЗА = "SQLSERVER", "OlapFactory"
ФАКТ = "Факти"


def g():
    return str(uuid.uuid4())


# Безопасный без кавычек идентификатор: только латиница, кириллица (вкл. і/ї/є/ґ), цифры и _.
# Намеренно УЖЕ, чем \w: \w в Python пропускает модификаторы вроде U+02BC (ʼ) и прочую
# не-ASCII пунктуацию, а как их разберёт парсер TMDL — не проверено ни разу.
БЕЗ_КАВЫЧЕК = re.compile(r"[A-Za-zЀ-ӿ0-9_]+")


def ид(имя):
    """TMDL-идентификатор.

    🔴 Два разных правила, оба выучены дорого:
    1) имя с пробелом/дефисом/любым несловесным символом ОБЯЗАНО быть в одинарных кавычках,
       иначе парсер валится `InvalidLineType: Непредвиденный тип строки: Other!`
       (проверено 02.08.2026 на `column Назва місяця`);
    2) одинарная кавычка ВНУТРИ имени обязана быть удвоена — иначе кавычки закрываются
       на середине имени и ломается тот же парсер. Ловушка не теоретическая: украинское
       «Підрозділ-об'єкт» даёт `column 'Підрозділ-об'єкт'` (найдено 03.08.2026 до передачи
       пользователю). В именах лучше вообще использовать типографский апостроф U+02BC (ʼ) —
       он не кавычка; удвоение здесь как страховка на случай ASCII-апострофа.
    Кириллица сама по себе кавычек не требует."""
    if БЕЗ_КАВЫЧЕК.fullmatch(имя):
        return имя
    return "'" + имя.replace("'", "''") + "'"


def кол(имя, тип, источник, скрыт=False, формат=None, свернуть="none",
        сортировка=None, ключ=False, аннотации=()):
    s = [f"\tcolumn {ид(имя)}", f"\t\tdataType: {тип}"]
    if ключ:                       # «пометить как таблицу дат» = isUnique + isKey на колонке даты
        s += ["\t\tisUnique", "\t\tisKey"]
    if скрыт:
        s.append("\t\tisHidden")
    if формат:
        s.append(f"\t\tformatString: {формат}")
    s += [f"\t\tlineageTag: {g()}", f"\t\tsummarizeBy: {свернуть}", f"\t\tsourceColumn: {источник}"]
    if сортировка:                 # без него месяцы и дни недели сортируются по алфавиту
        s.append(f"\t\tsortByColumn: {ид(сортировка)}")
    s += ["", "\t\tannotation SummarizationSetBy = Automatic", ""]
    for а in аннотации:
        s += [f"\t\tannotation {а}", ""]
    return "\n".join(s)


def партиция(таблица, колонки, объект):
    список = ", ".join(колонки)
    m = (f'\tpartition {ид(таблица)} = m\n'
         f'\t\tmode: import\n'
         f'\t\tsource =\n'
         f'\t\t\t\tlet\n'
         f'\t\t\t\t    Джерело = Sql.Database("{СЕРВЕР}", "{БАЗА}", '
         f'[Query="SELECT {список} FROM dbo.{объект}"])\n'
         f'\t\t\t\tin\n'
         f'\t\t\t\t    Джерело\n')
    return m


def таблица(имя, колонки_tmdl, партиция_tmdl, меры_tmdl="", свойства=()):
    тело = f"table {ид(имя)}\n\tlineageTag: {g()}\n"
    for с in свойства:
        тело += f"\t{с}\n"
    тело += "\n"
    тело += меры_tmdl
    тело += "\n".join(колонки_tmdl)
    тело += "\n" + партиция_tmdl
    тело += "\n\tannotation PBI_ResultType = Table\n"
    return тело


# ------------------------------------------------------------------ меры факта
# 🔴 Имя таблицы с НЕ-ASCII символами в DAX обязано быть в одинарных кавычках:
#    SUM(Факти[X]) -> синтаксическая ошибка «недопустимый токен ... Ф»;
#    SUM('Факти'[X]) -> корректно. Проверено на живой модели 02.08.2026.
#    Ссылки на другие меры ([Имя]) кавычек не требуют.
МЕРЫ = [
    ("План, год",                    "CALCULATE(SUM('Факти'[PlanHours]), REMOVEFILTERS('Calendar'))", "#,##0.00"),
    ("Факт, год",                    "SUM('Факти'[FactHours])", "#,##0.00"),
    ("Виконання, год",               "SUM('Факти'[EarnedHours])", "#,##0.00"),
    ("Відхилення, год",              "[Факт, год] - [План, год]", "#,##0.00"),
    ("% вик. (год)",                 "DIVIDE([Факт, год], [План, год]) * 100", "#,##0.0"),
    ("Відхилення табель-виконання",  "[Факт, год] - [Виконання, год]", "#,##0.00"),
    ("% виконання",                  "DIVIDE([Виконання, год], [План, год]) * 100", "#,##0.0"),
    ("% виконання до факту",         "DIVIDE([Виконання, год], [Факт, год]) * 100", "#,##0.0"),
    ("План, к-сть",                  "CALCULATE(SUM('Факти'[PlanQty]), REMOVEFILTERS('Calendar'))", "#,##0.000"),
    ("Факт, к-сть",                  "SUM('Факти'[FactQty])", "#,##0.000"),
    ("Відхилення, к-сть",            "[Факт, к-сть] - [План, к-сть]", "#,##0.000"),
    ("% вик. (к-сть)",               "DIVIDE([Факт, к-сть], [План, к-сть]) * 100", "#,##0.0"),
    ("План, грн",                    "CALCULATE(SUM('Факти'[PlanUAH]), REMOVEFILTERS('Calendar'))", "#,##0.00"),
    ("Факт, грн",                    "SUM('Факти'[FactUAH])", "#,##0.00"),
    ("Відхилення, грн",              "[Факт, грн] - [План, грн]", "#,##0.00"),
    ("% вик. (грн)",                 "DIVIDE([Факт, грн], [План, грн]) * 100", "#,##0.0"),
    ("Ціна план, грн",               "DIVIDE([План, грн], [План, к-сть])", "#,##0.00"),
    ("Ціна факт, грн",               "DIVIDE([Факт, грн], [Факт, к-сть])", "#,##0.00"),
    ("Плановий розрахунок матеріалів, грн", "CALCULATE(SUM('Факти'[ETC_UAH]), REMOVEFILTERS('Calendar'))", "#,##0.00"),
    ("План на факт, грн",            "[Факт, грн] + [Плановий розрахунок матеріалів, грн]", "#,##0.00"),
    ("Прогнозне відхилення, грн",    "[План на факт, грн] - [План, грн]", "#,##0.00"),
    ("% відхилення",                 "DIVIDE([План на факт, грн], [План, грн]) * 100", "#,##0.0"),
]


def меры_tmdl():
    out = []
    for имя, dax, фмт in МЕРЫ:
        out.append(f"\tmeasure '{имя}' = {dax}\n"
                   f"\t\tformatString: {фмт}\n"
                   f"\t\tlineageTag: {g()}\n")
    return "\n".join(out) + "\n"


def main():
    корень = ВЫХОД / ИМЯ
    модель = ВЫХОД / f"{ИМЯ}.SemanticModel"
    отчет = ВЫХОД / f"{ИМЯ}.Report"
    for d in (модель, отчет):
        if d.exists():
            shutil.rmtree(d)
    (модель / "definition" / "tables").mkdir(parents=True)
    (отчет / "definition" / "pages").mkdir(parents=True)

    # -------------------------------------------------- .pbip
    (ВЫХОД / f"{ИМЯ}.pbip").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{ИМЯ}.Report"}}],
        "settings": {"enableAutoRecovery": True},
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    for папка, тип in ((модель, "SemanticModel"), (отчет, "Report")):
        (папка / ".platform").write_text(json.dumps({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
            "metadata": {"type": тип, "displayName": ИМЯ},
            "config": {"version": "2.0", "logicalId": g()},
        }, ensure_ascii=False, indent=2), encoding="utf-8")

    (модель / "definition.pbism").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2", "settings": {},
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    (отчет / "definition.pbir").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0",
        "datasetReference": {"byPath": {"path": f"../{ИМЯ}.SemanticModel"}},
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    (модель / "definition" / "database.tmdl").write_text(
        "database\n\tcompatibilityLevel: 1606\n\n", encoding="utf-8")

    # -------------------------------------------------- таблицы
    таблицы = {}

    таблицы[ФАКТ] = таблица(ФАКТ, [
        кол("Period", "dateTime", "Period", формат="General Date"),
        кол("Period_Month", "dateTime", "Period_Month", скрыт=True, формат="General Date"),
        кол("Organization_ID", "string", "Organization_ID", скрыт=True),
        кол("Department_ID", "string", "Department_ID", скрыт=True),
        кол("ExecutorDept_ID", "string", "ExecutorDept_ID", скрыт=True),
        кол("RowType", "string", "RowType", скрыт=True),
        кол("Блок", "string", "Block"),
        кол("Категорія", "string", "Category"),
        кол("Stage_ID", "string", "Stage_ID", скрыт=True),
        кол("Work_ID", "string", "Work_ID", скрыт=True),
        кол("CommonName_ID", "string", "CommonName_ID", скрыт=True),
        кол("Аналітика", "string", "Analytics_Name"),
        кол("Document_ID", "string", "Document_ID", скрыт=True),
        кол("ObjectDept_ID", "string", "ObjectDept_ID", скрыт=True),
        кол("Unit_ID", "string", "Unit_ID", скрыт=True),
        # реквизит «НоменклатураНаименование» регистра: снимок имени на момент проведения,
        # заполнен только у факта закупок (1287 строк). В эталонном отчёте титул русский —
        # «Наименование номенклатуры»; модель украиноязычная, поэтому имя переведено.
        кол("Найменування номенклатури", "string", "ItemName"),
        кол("PlanHours", "double", "PlanHours", скрыт=True, свернуть="sum"),
        кол("FactHours", "double", "FactHours", скрыт=True, свернуть="sum"),
        кол("EarnedHours", "double", "EarnedHours", скрыт=True, свернуть="sum"),
        кол("PlanQty", "double", "PlanQty", скрыт=True, свернуть="sum"),
        кол("FactQty", "double", "FactQty", скрыт=True, свернуть="sum"),
        кол("PlanUAH", "double", "PlanUAH", скрыт=True, свернуть="sum"),
        кол("FactUAH", "double", "FactUAH", скрыт=True, свернуть="sum"),
        кол("ETC_UAH", "double", "ETC_UAH", скрыт=True, свернуть="sum"),
    ], партиция(ФАКТ, [
        "Period", "Period_Month", "Organization_ID", "Department_ID", "ExecutorDept_ID",
        "RowType", "Block", "Category", "Stage_ID", "Work_ID", "CommonName_ID",
        "Analytics_Name", "Document_ID", "ObjectDept_ID", "Unit_ID", "ItemName",
        "PlanHours", "FactHours", "EarnedHours",
        "PlanQty", "FactQty", "PlanUAH", "FactUAH", "ETC_UAH"], "vw_Fact_PlanFact"),
        меры_tmdl())

    # 🔴 Три роли одного справочника подразделений (role-playing dimensions).
    # Power BI допускает между парой таблиц лишь ОДИН активный путь, поэтому «підрозділ»,
    # «підрозділ-виконавець» и «підрозділ-обʼєкт» — три отдельные таблицы поверх одного и того же
    # представления, у каждой своя АКТИВНАЯ связь к своему ключу факта. Альтернатива
    # (одна таблица + USERELATIONSHIP) потребовала бы дублировать все 22 меры на каждый разрез.
    def подразделения(таблица_имя, поле_имя):
        return таблица(таблица_имя, [
            кол("Department_ID", "string", "Department_ID", скрыт=True),
            кол(поле_имя, "string", "Department_Name"),
            кол("Код", "string", "Department_Code", скрыт=True),
            кол("Рівень 1", "string", "Level1"),
            кол("Рівень 2", "string", "Level2"),
            кол("Рівень 3", "string", "Level3"),
            кол("Рівень 4", "string", "Level4"),
            кол("Шлях", "string", "Hierarchy_Path", скрыт=True),
        ], партиция(таблица_имя, ["Department_ID", "Department_Name", "Department_Code",
                                  "Level1", "Level2", "Level3", "Level4", "Hierarchy_Path"],
                    "Dim_Departments_Tree"))

    # «хозяин» строки — подразделение структуры себестоимости, заполнено у ВСЕХ строк
    таблицы["Підрозділи"] = подразделения("Підрозділи", "Підрозділ")
    # кто фактически отработал: подразделение табеля и ВРМ, включая вложенные «№1»/«№2».
    # Заполнено у факта (1636) и виконання (168); у плана и ETC пусто — как в эталонном отчёте.
    таблицы["ПідрозділиВиконавці"] = подразделения("ПідрозділиВиконавці", "Підрозділ-виконавець")
    # объект 15м/30м: реквизит карточки СС. Заполнен у плана (1069) и виконання (168);
    # у факта пусто («котёл» на итоге подразделения), у ETC пусто по решению пользователя.
    # 🔴 апостроф — типографский U+02BC (ʼ), НЕ ASCII-кавычка: внутри TMDL-идентификатора
    # ASCII-апостроф закрывает кавычку на середине имени. Заодно это верная украинская типографика.
    таблицы["ПідрозділиОбʼєкти"] = подразделения("ПідрозділиОбʼєкти", "Підрозділ-обʼєкт")

    таблицы["Етапи"] = таблица("Етапи", [
        кол("Stage_ID", "string", "Stage_ID", скрыт=True),
        кол("Етап", "string", "Stage_Name"),
    ], партиция("Етапи", ["Stage_ID", "Stage_Name"], "Dim_Stages"))

    таблицы["Роботи"] = таблица("Роботи", [
        кол("Item_ID", "string", "Item_ID", скрыт=True),
        кол("Робота", "string", "Item_Name"),
        кол("Код", "string", "Item_Code", скрыт=True),
    ], партиция("Роботи", ["Item_ID", "Item_Name", "Item_Code"], "Dim_Items"))

    таблицы["ЗагальніНазви"] = таблица("ЗагальніНазви", [
        кол("CommonName_ID", "string", "CommonName_ID", скрыт=True),
        кол("Загальна назва", "string", "CommonName"),
    ], партиция("ЗагальніНазви", ["CommonName_ID", "CommonName"], "Dim_CommonNames"))

    таблицы["Одиниці"] = таблица("Одиниці", [
        кол("Unit_ID", "string", "Unit_ID", скрыт=True),
        кол("Одиниця", "string", "Unit_Name"),
    ], партиция("Одиниці", ["Unit_ID", "Unit_Name"], "Dim_Units"))

    таблицы["Організації"] = таблица("Організації", [
        кол("Organization_ID", "string", "Organization_ID", скрыт=True),
        кол("Організація", "string", "Organization_Name"),
    ], партиция("Організації", ["Organization_ID", "Organization_Name"], "Dim_Organizations"))

    таблицы["ТипРядка"] = таблица("ТипРядка", [
        кол("RowType", "string", "RowType"),
        кол("Сортування", "int64", "RowType_Sort", скрыт=True),
        кол("Група", "string", "RowType_Group"),
    ], партиция("ТипРядка", ["RowType", "RowType_Sort", "RowType_Group"], "Dim_RowTypes"))

    # 🔴 Документи — расшифровка каждой цифры до первички 1С (табель / приход / ВРМ /
    # карточка СС). Строки ETC документа не имеют by design: это расчётный остаток плана,
    # а не проводка. Раскладка проверена: Виконання -> 27 ВРМ, План -> 5 карточек СС,
    # Факт -> 356 приходов + 20 табелей, 0 FK-сирот.
    таблицы["Документи"] = таблица("Документи", [
        кол("Document_ID", "string", "Document_ID", скрыт=True),
        кол("Документ", "string", "Document_Presentation"),
        кол("Тип документа", "string", "Document_Type"),
        кол("Номер", "string", "Document_Number"),
        кол("Дата документа", "dateTime", "Document_Date", формат="General Date"),
    ], партиция("Документи", ["Document_ID", "Document_Presentation", "Document_Type",
                              "Document_Number", "Document_Date"], "Dim_Documents"))

    # 🔴 Calendar — зеркало эталона PowerBi/Industrial/PL.SemanticModel/…/Calendar.tmdl:
    # те же 43 колонки, те же имена, типы, форматы; dataCategory: Time; isUnique+isKey
    # на date_ (это и есть «пометить как таблицу дат») и 17 sortByColumn — без них месяцы
    # и дни недели сортируются по алфавиту.
    # Отличий от эталона ровно два, оба намеренные:
    #   * источник — OlapFactory (у эталона OlapBASERP) и читается NATIVE QUERY с явным
    #     списком колонок: Sql.Database кеширует схему и при [Schema=,Item=] новые колонки
    #     не появятся даже после Full Refresh;
    #   * три чистых ключа сортировки скрыты (isHidden) — в отчёте они не нужны, а на
    #     сортировку по ним скрытие не влияет.
    ФБУЛ = '"""TRUE"";""TRUE"";""FALSE"""'
    ДАТАМЕС = ("UnderlyingDateTimeDataType = Date",)

    def ц(имя, сорт=None, скрыт=False):      # целое
        return кол(имя, "int64", имя, скрыт=скрыт, формат="0", сортировка=сорт)

    def т(имя, сорт=None):                   # текст
        return кол(имя, "string", имя, сортировка=сорт)

    def б(имя):                              # булево
        return кол(имя, "boolean", имя, формат=ФБУЛ)

    def дм(имя):                             # начало месяца/квартала/года
        return кол(имя, "dateTime", имя, формат="Long Date", аннотации=ДАТАМЕС)

    календарь = [
        кол("date_", "dateTime", "date_", формат="General Date", ключ=True),
        ц("year_"), ц("quarter_"), т("Квартал", "quarter_"), ц("month_"), ц("day_"),
        ц("year_day"), ц("week_"), ц("week_day"),
        т("month_name", "month_"), т("month_name_en", "month_"), т("month_name_ua", "month_"),
        т("month_name_short", "month_"), т("month_name_short_en", "month_"),
        т("week_day_str", "week_day"), т("week_day_str_en", "week_day"),
        т("week_day_str_ua", "week_day"), т("week_day_short", "week_day"),
        т("week_day_short_en", "week_day"),
        т("year_week", "Year_Week_Sort"), т("year_month", "Year_Month_Sort"),
        т("Year_Quarter", "Year_Quarter_Sort"), т("year_quarter_en", "Year_Quarter_Sort"),
        т("date_str"), т("date_full_str"), т("date_full_str_en"),
        ц("month_week"), ц("month_decade"), т("month_last_day"), т("year_last_day"),
        ц("days_in_month"), ц("days_in_year"), б("weekend"),
        дм("Period_Month"), дм("Period_Quarter"), дм("Period_Year"),
        ц("Year_Month_Sort", скрыт=True), ц("Year_Quarter_Sort", скрыт=True),
        ц("Year_Week_Sort", скрыт=True),
        б("Is_Working_Day"), ц("Days_To_Month_End"),
        т("month_name_short_ua", "month_"), т("year_month_short_ua", "Year_Month_Sort"),
    ]
    источники = [re.search(r"sourceColumn: (.+)", c).group(1) for c in календарь]
    таблицы["Calendar"] = таблица("Calendar", календарь,
                                  партиция("Calendar", источники, "Calendar"),
                                  свойства=("dataCategory: Time",))

    for имя, тело in таблицы.items():
        (модель / "definition" / "tables" / f"{имя}.tmdl").write_text(тело, encoding="utf-8")

    # -------------------------------------------------- связи
    связи = [
        ("fact_department", f"{ФАКТ}.Department_ID", "Підрозділи.Department_ID", True),
        # обе роли — АКТИВНЫЕ: ведут в разные таблицы, второго пути к «Підрозділи» не возникает
        ("fact_executor_dept", f"{ФАКТ}.ExecutorDept_ID", "ПідрозділиВиконавці.Department_ID", True),
        ("fact_object_dept", f"{ФАКТ}.ObjectDept_ID", "ПідрозділиОбʼєкти.Department_ID", True),
        ("fact_organization", f"{ФАКТ}.Organization_ID", "Організації.Organization_ID", True),
        ("fact_stage", f"{ФАКТ}.Stage_ID", "Етапи.Stage_ID", True),
        ("fact_work", f"{ФАКТ}.Work_ID", "Роботи.Item_ID", True),
        ("fact_common_name", f"{ФАКТ}.CommonName_ID", "ЗагальніНазви.CommonName_ID", True),
        ("fact_unit", f"{ФАКТ}.Unit_ID", "Одиниці.Unit_ID", True),
        ("fact_row_type", f"{ФАКТ}.RowType", "ТипРядка.RowType", True),
        ("fact_document", f"{ФАКТ}.Document_ID", "Документи.Document_ID", True),
        ("fact_calendar", f"{ФАКТ}.Period", "Calendar.date_", True),
    ]
    # ссылка «Таблица.Колонка»: имя таблицы подчиняется тем же правилам кавычек, что и объявление,
    # иначе связь на таблицу с не-словесным символом в имени не разберётся
    def ссылка(путь):
        т, _, к = путь.rpartition(".")
        return f"{ид(т)}.{ид(к)}"

    блоки = []
    for имя, откуда, куда, активна in связи:
        b = f"relationship {имя}\n"
        if not активна:
            b += "\tisActive: false\n"
        b += f"\tfromColumn: {ссылка(откуда)}\n\ttoColumn: {ссылка(куда)}\n"
        блоки.append(b)
    (модель / "definition" / "relationships.tmdl").write_text("\n".join(блоки), encoding="utf-8")

    # -------------------------------------------------- model.tmdl
    порядок = list(таблицы)
    m = ("model Model\n\tculture: uk-UA\n\tdefaultPowerBIDataSourceVersion: powerBI_V3\n"
         "\tsourceQueryCulture: uk-UA\n\tdataAccessOptions\n\t\tlegacyRedirects\n"
         "\t\treturnErrorValuesAsNull\n\n"
         "annotation __PBI_TimeIntelligenceEnabled = 0\n\n"
         f"annotation PBI_QueryOrder = {json.dumps(порядок, ensure_ascii=False)}\n\n")
    m += "".join(f"ref table {ид(t)}\n" for t in порядок)
    (модель / "definition" / "model.tmdl").write_text(m, encoding="utf-8")

    # -------------------------------------------------- макет модели (Desktop ждёт его рядом)
    (модель / "diagramLayout.json").write_text(json.dumps({
        "version": "1.1.0", "diagrams": [{
            "ordinal": 0, "scrollPosition": {"x": 0, "y": 0}, "nodes": [],
            "name": "All tables", "zoomValue": 100, "pinKeyFieldsToTop": False,
            "showExtraHeaderInfo": False, "hideKeyFieldsWhenCollapsed": False,
            "tablesLocked": False,
        }], "selectedDiagram": "All tables", "defaultDiagram": "All tables",
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    # -------------------------------------------------- отчёт
    # 🔴 Структура повторяет заведомо рабочий PowerBi/Industrial/PL.Report:
    #    definition/version.json ОБЯЗАТЕЛЕН — без него Desktop падает
    #    «Cannot find file 'version.json'» ещё до чтения модели;
    #    тема, на которую ссылается report.json, должна физически лежать в StaticResources.
    (отчет / "definition" / "version.json").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
        "version": "2.0.0",
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    тема_src = Path(r"C:/Configuration_downloads/BASERP25/_Rarzrabotki/PowerBi/Industrial/PL.Report/StaticResources/SharedResources/BaseThemes/CY26SU04.json")
    if тема_src.exists():
        тема_dst = отчет / "StaticResources" / "SharedResources" / "BaseThemes" / "CY26SU04.json"
        тема_dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(тема_src, тема_dst)
        тема_ок = True
    else:
        тема_ок = False

    report_json = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.2.0/schema.json",
        "settings": {
            "useStylableVisualContainerHeader": True,
            "exportDataMode": "AllowSummarized",
            "defaultDrillFilterOtherVisuals": True,
            "allowChangeFilterTypes": True,
            "useEnhancedTooltips": True,
            "useDefaultAggregateDisplayName": True,
        },
    }
    if тема_ок:
        report_json["themeCollection"] = {"baseTheme": {
            "name": "CY26SU04",
            "reportVersionAtImport": {"visual": "2.8.0", "report": "3.2.0", "page": "2.3.1"},
            "type": "SharedResources"}}
        report_json["resourcePackages"] = [{
            "name": "SharedResources", "type": "SharedResources",
            "items": [{"name": "CY26SU04", "path": "BaseThemes/CY26SU04.json", "type": "BaseTheme"}]}]
    (отчет / "definition" / "report.json").write_text(
        json.dumps(report_json, ensure_ascii=False, indent=2), encoding="utf-8")

    стр = g().replace("-", "")[:20]
    (отчет / "definition" / "pages" / "pages.json").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
        "pageOrder": [стр], "activePageName": стр,
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    страница = отчет / "definition" / "pages" / стр
    страница.mkdir(parents=True, exist_ok=True)
    (страница / "page.json").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json",
        "name": стр, "displayName": "План-факт виробництва",
        "displayOption": "FitToWidth", "height": 900, "width": 1400,
        "pageBinding": {"name": g().replace("-", "")[:20], "type": "Default",
                        "parameters": [], "acceptsFilterContext": "None"},
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[OK] PBIP собран: {ВЫХОД / (ИМЯ + '.pbip')}")
    print(f"     таблиц: {len(таблицы)}, мер: {len(МЕРЫ)}, связей: {len(связи)}")
    for f in sorted(ВЫХОД.rglob("*")):
        if f.is_file() and "Py_Olap" not in str(f):
            print("     ", f.relative_to(ВЫХОД))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
