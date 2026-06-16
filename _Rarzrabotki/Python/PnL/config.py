"""Central config for PnL import pipeline."""
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
DATA_DIR = PROJECT_ROOT / "data"
JSON_DIR = DATA_DIR / "json"
REPORTS_DIR = DATA_DIR / "reports"

EXCEL_FILES = [
    # === 2024 ГОД (12 месяцев из одного декабрьского Excel 2024; колонки col=14..69 step=5) ===
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-01-31", "month_header": "2024-01-01", "label": "Январь 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-02-29", "month_header": "2024-02-01", "label": "Февраль 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-03-31", "month_header": "2024-03-01", "label": "Март 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-04-30", "month_header": "2024-04-01", "label": "Апрель 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-05-31", "month_header": "2024-05-01", "label": "Май 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-06-30", "month_header": "2024-06-01", "label": "Июнь 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-07-31", "month_header": "2024-07-01", "label": "Июль 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-08-31", "month_header": "2024-08-01", "label": "Август 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-09-30", "month_header": "2024-09-01", "label": "Сентябрь 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-10-31", "month_header": "2024-10-01", "label": "Октябрь 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-11-30", "month_header": "2024-11-01", "label": "Ноябрь 2024"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь 2024\!PL по компании Декабрь 2024.xlsx",
     "period": "2024-12-31", "month_header": "2024-12-01", "label": "Декабрь 2024"},
    # === 2025 ГОД (11 месяцев из одного декабрьского Excel; колонки col=19..69 step=5) ===
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx",
     "period": "2025-01-31", "month_header": "2025-01-01", "label": "Январь 2025"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx",
     "period": "2025-02-28", "month_header": "2025-02-01", "label": "Февраль 2025"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx",
     "period": "2025-03-31", "month_header": "2025-03-01", "label": "Март 2025"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx",
     "period": "2025-04-30", "month_header": "2025-04-01", "label": "Апрель 2025"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx",
     "period": "2025-05-31", "month_header": "2025-05-01", "label": "Май 2025"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx",
     "period": "2025-06-30", "month_header": "2025-06-01", "label": "Июнь 2025"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx",
     "period": "2025-07-31", "month_header": "2025-07-01", "label": "Июль 2025"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx",
     "period": "2025-08-31", "month_header": "2025-08-01", "label": "Август 2025"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx",
     "period": "2025-09-30", "month_header": "2025-09-01", "label": "Сентябрь 2025"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx",
     "period": "2025-10-31", "month_header": "2025-10-01", "label": "Октябрь 2025"},
    {"path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx",
     "period": "2025-11-30", "month_header": "2025-11-01", "label": "Ноябрь 2025"},
    {
        "path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Декабрь_25\!PL по компании Декабрь 2025.xlsx",
        "period": "2025-12-31",
        "month_header": "2025-12-01",
        "label": "Декабрь 2025",
    },
    {
        "path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Январь_26\!PL по компании Январь 2026.xlsx",
        "period": "2026-01-31",
        "month_header": "2026-01-01",
        "label": "Январь 2026",
    },
    {
        "path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Лютий_26\!PL по компании Лютий 2026.xlsx",
        "period": "2026-02-28",
        "month_header": "2026-02-01",
        "label": "Лютий 2026",
    },
    {
        "path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Березень_26\!PL по компании Березень 2026.xlsx",
        "period": "2026-03-31",
        "month_header": "2026-03-01",
        "label": "Березень 2026",
    },
    {
        "path": r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PL\Апрель_26\!PL по компании Квітень 2026.xlsx",
        "period": "2026-04-30",
        "month_header": "2026-04-01",
        "label": "Квітень 2026",
    },
]

CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

ORGANIZATION_NAME = 'ТОВ "ІНДАСТРІАЛБУД"'

SUMMARY_SHEET_PATTERNS = [
    "PL_", "PL ", "МК ", "МД ", "MAN ", "Форд ", "Метрики", "Индекс",
    "ЦО_", "Лист", "расчет ", "ОС", "РАЗНИЦА",
    # "Техника" убрана из фильтра — решение финансиста 2026-04-17:
    # лист Excel "Техника" = подразделение 1С "Логистика" (корневое),
    # сворачивает MAN-ы и Экспедирование в один PnL-лист.
    "Экспедирование", "Цех ЗП",
    # "Логистика" оставлена в SUMMARY (обычно сводный лист), но если встречается
    # отдельный лист с данными по подразделению "Логистика" — замапится через 07.
    "Логистика",
]

# Ручные override для маппинга лист Excel → подразделение 1С.
# Применяется в 07_match_sheets_to_struct ПОСЛЕ fuzzy-алгоритма.
# Добавляй сюда пары, когда fuzzy-score < FUZZY_AUTO_THRESHOLD или когда имя листа Excel
# сознательно не совпадает с 1С-подразделением (напр. агрегация).
MANUAL_SHEET_TO_STRUCT_OVERRIDES = {
    # === Обновлено 2026-04-18 по эталону Документ.А_РасшифровкаЛистов №000000002 (финансист) ===
    # Техника — свод по направлению "Спецтехника" (НЕ Логистика!).
    "Техника": {
        "struct_uuid": "cdf975d7-ca29-11f0-a2e2-c54425f51b91",
        "struct_name": "Спецтехника",
        "direction_uuid": "9d021b6e-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Спецтехника",
        "include_children": True,
        "reason": "Решение финансиста (док №2): лист Техника = свод по направлению Спецтехника",
    },
    # PL_Логистика — новый сводный лист, появился в январе 2026.
    "PL_Логистика": {
        "struct_uuid": "950c32e4-ca27-11f0-a2e2-c54425f51b91",
        "struct_name": "Логистика",
        "direction_uuid": "9d2c848d-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Логистика",
        "include_children": True,
        "reason": "Сводный лист по направлению Логистика (добавлен финансистом в док №2)",
    },
    # fuzzy-score 80 % ниже AUTO_THRESHOLD=85, поэтому 07-алгоритм не матчит автоматически.
    # Имя в 1С — "МД ПРООН Черкаси  ДСНС" (два пробела!).
    "ПРОООН Черкаси ДСНС": {
        "struct_uuid": "1ed6ec43-f859-11f0-8105-00155dce3d04",
        "struct_name": "МД ПРООН Черкаси  ДСНС",
        "reason": "fuzzy-score ниже порога + опечатка в имени листа Excel (ПРОООН vs ПРООН)",
    },
    # === Добавлено 2026-04-18 по правкам финансиста в Документ.А_РасшифровкаЛистов №000000002 ===
    # Сводные листы ЦО / Строительство / Производство.
    "PL_ЦО": {
        "struct_uuid": "19c7dbef-ca2d-11f0-a2e2-c54425f51b91",
        "struct_name": "ЦО",
        "direction_uuid": "9d021b6f-1ae7-11f0-80dc-00155d235309",
        "direction_name": "ЦО",
        "include_children": True,
        "reason": "Сводный лист ЦО = корневое подразделение ЦО",
    },
    "PL Строительство Свод": {
        "struct_uuid": "98e5f138-ca29-11f0-a2e2-c54425f51b91",
        "struct_name": "Строительство",
        "direction_uuid": "9d021b84-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Строительство",
        "include_children": True,
        "reason": "Сводный лист Строительство = корневое подразделение Строительство",
    },
    "PL_ЦО_Строительство": {
        "struct_uuid": "98e5f138-ca29-11f0-a2e2-c54425f51b91",
        "struct_name": "Строительство",
        "direction_uuid": "9d021b84-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Строительство",
        "include_children": False,
        "reason": "ЦО по Строительству (док №2: не свод, просто привязка к направлению)",
    },
    "PL_Производство_СВОД": {
        "struct_uuid": "65c4e837-ca29-11f0-a2e2-c54425f51b91",
        "struct_name": "Производство",
        "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Производство",
        "include_children": True,
        "reason": "Сводный лист Производство = корневое подразделение Производство",
    },
    "PL_ЦО_Производство": {
        "struct_uuid": "65c4e837-ca29-11f0-a2e2-c54425f51b91",
        "struct_name": "Производство",
        "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Производство",
        "include_children": False,
        "reason": "ЦО по Производству (док №2: не свод, просто привязка к направлению)",
    },
    # Межфилиальные листы.
    "МК Глобино": {
        "struct_uuid": "9d2c918b-1ae7-11f0-80dc-00155d235309",
        "struct_name": "МК Глобино",
        "reason": "Точное соответствие МК Глобино (не попало в exact из-за фильтра 01_raw_sheets?)",
    },
    "МД ООН 2025": {
        "struct_uuid": "7bd47f9f-50fa-11f0-80f3-00155dce3d04",
        "struct_name": "МД ООН 2025",
        "reason": "Точное соответствие МД ООН 2025",
    },
    "МД ВООЗ 2025": {
        "struct_uuid": "92a51c8e-ab51-11f0-8100-00155dce3d04",
        "struct_name": "МД ВООЗ  2025",
        "reason": "В 1С имя с двумя пробелами: 'МД ВООЗ  2025'",
    },
    # МК Чорноморск (добавлен в док №3, Лютий 2026) — МК-филиал, направление Производство.
    # Отдельно от листа "Черноморск" (Чорноморськ, Строительство).
    "МК Чорноморск": {
        "struct_uuid": "dcc4e9ba-fda0-11f0-8105-00155dce3d04",
        "struct_name": "МК Чорноморськ",
        "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Производство",
        "include_children": False,
        "reason": "МК Чорноморск = МК Чорноморськ (Производство). Эталон: док №3 Лютий 2026",
    },
    # МД МХП ОРІЛЬ — новый МД-филиал, появился в апреле 2026. Направление Производство.
    "МД МХП ОРІЛЬ": {
        "struct_uuid": "1f60d4b1-26b2-11f1-810c-00155dce3d04",
        "struct_name": "МД МХП ОРІЛЬ",
        "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Производство",
        "include_children": False,
        "reason": "Новый МД-филиал в Квітень 2026 (код 00-001020). Производство.",
    },
    # === 2024 ГОД — филиалы, закрытые/изменённые к 2025 (добавлено 2026-05-21) ===
    "PL_Староконстантинов": {
        "struct_uuid": "9db438f5-1ae7-11f0-80dc-00155d235309",
        "struct_name": "Староконстантинов",
        "direction_uuid": "9d021b84-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Строительство",
        "include_children": True,
        "reason": "Сводный лист PL_Староконстантинов = подразделение Строительства",
    },
    "Молоко Вітчизни": {
        "struct_uuid": "9d2c90df-1ae7-11f0-80dc-00155d235309",
        "struct_name": "Молоко Вітчизни",
        "direction_uuid": "9d021b84-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Строительство",
        "include_children": False,
        "reason": "Подразделение 2024 года, Строительство",
    },
    "КМД Путровка": {
        "struct_uuid": "9d2c911d-1ae7-11f0-80dc-00155d235309",
        "struct_name": "КМД Путровка",
        "direction_uuid": "9d021b71-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Закрытые обьекты",
        "include_children": False,
        "reason": "Объект закрыт к 2025",
    },
    "Турбина-3": {
        "struct_uuid": "9d2c90b2-1ae7-11f0-80dc-00155d235309",
        "struct_name": "Турбина 3",
        "direction_uuid": "9d2c84a1-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Девелопмент",
        "include_children": False,
        "reason": "В Excel 'Турбина-3', в 1С — 'Турбина 3' без дефиса",
    },
    "Нежин 3": {
        "struct_uuid": "9d2c8fc5-1ae7-11f0-80dc-00155d235309",
        "struct_name": "Нежин 3 ",  # trailing space в 1С!
        "direction_uuid": "9d021b71-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Закрытые обьекты",
        "include_children": False,
        "reason": "Объект закрыт; в 1С имя с trailing space",
    },
    "Глобино": {
        "struct_uuid": "9d2c90ca-1ae7-11f0-80dc-00155d235309",
        "struct_name": "Глобино",
        "direction_uuid": "9d021b71-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Закрытые обьекты",
        "include_children": False,
        "reason": "Объект 'Глобино' закрыт к 2025 (отдельный от 'Глобино-2' и 'МК Глобино')",
    },
    "Приднепровский_металл": {
        "struct_uuid": "9dc455b5-1ae7-11f0-80dc-00155d235309",
        "struct_name": "Приднепровский-металл",
        "direction_uuid": "9d021b71-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Закрытые обьекты",
        "include_children": False,
        "reason": "В Excel underscore, в 1С — дефис",
    },
    "Солома_Элеватор": {
        "struct_uuid": "9d2c8c54-1ae7-11f0-80dc-00155d235309",
        "struct_name": "Солома Элеватор",
        "include_children": False,
        "reason": "В Excel underscore, в 1С — пробел; направление не задано",
    },
    "МД Первомайск ПА": {
        "struct_uuid": "9d2c9319-1ae7-11f0-80dc-00155d235309",
        "struct_name": "МД Первомайск ПА",
        "direction_uuid": "9d021b71-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Закрытые обьекты",
        "include_children": False,
        "reason": "Точное совпадение, объект закрыт к 2025",
    },
    "НГУ": {
        "struct_uuid": "9d2c91e6-1ae7-11f0-80dc-00155d235309",
        "struct_name": "НГУ",
        "include_children": False,
        "reason": "Точное совпадение, направление не задано",
    },
    "МД ООН": {
        "struct_uuid": "9d2c8fb1-1ae7-11f0-80dc-00155d235309",
        "struct_name": "МД ООН",
        "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Производство",
        "include_children": False,
        "reason": "Отдельный от 'МД ООН 2025'",
    },
    "МД ВООЗ": {
        "struct_uuid": "9d2c90c9-1ae7-11f0-80dc-00155d235309",
        "struct_name": "МД ВООЗ",
        "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Производство",
        "include_children": False,
        "reason": "Отдельный от 'МД ВООЗ  2025'",
    },
    "ЕМС": {
        "struct_uuid": "9d2c92f3-1ae7-11f0-80dc-00155d235309",
        "struct_name": "ЕМС малий",
        "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Производство",
        "include_children": False,
        "reason": "В Excel короткое 'ЕМС', в 1С — 'ЕМС малий'",
    },
    "Форд Тягач": {
        "struct_uuid": "9dc44b0c-1ae7-11f0-80dc-00155d235309",
        "struct_name": "Тягач Форд  1842Т",
        "include_children": False,
        "reason": "В 1С 'Тягач Форд  1842Т' с двумя пробелами",
    },
    "MAN Blue KA1783HI": {
        "struct_uuid": "9de72248-1ae7-11f0-80dc-00155d235309",
        "struct_name": "MAN blue № 1 KA1783HI",
        "include_children": False,
        "reason": "В Excel 'MAN Blue KA1783HI', в 1С — 'MAN blue № 1 KA1783HI'",
    },
    "MAN green KA2790IE": {
        "struct_uuid": "9de7226d-1ae7-11f0-80dc-00155d235309",
        "struct_name": "MAN green KA2790IE",
        "include_children": False,
        "reason": "Точное совпадение",
    },
    # МК-филиалы с точным совпадением имени. Направление — Производство (из А_НаправлениеДеятельности подразделения).
    "МК Астарта.Тищенки": {
        "struct_uuid": "a4e02be3-a36b-11f0-8100-00155dce3d04",
        "struct_name": "МК Астарта.Тищенки",
        "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Производство",
        "include_children": False,
        "reason": "МК Астарта.Тищенки = подразделение с тем же именем, Производство (эталон док №1)",
    },
    "МК МХП Катеринополь": {
        "struct_uuid": "9d2c9554-1ae7-11f0-80dc-00155d235309",
        "struct_name": "МК МХП Катеринополь",
        "direction_uuid": "9d2c848e-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Производство",
        "include_children": False,
        "reason": "МК МХП Катеринополь = подразделение с тем же именем, Производство (эталон док №1)",
    },
    # Логистика (детализация).
    "ЦО_логистика": {
        "struct_uuid": "950c32e4-ca27-11f0-a2e2-c54425f51b91",
        "struct_name": "Логистика",
        "direction_uuid": "9d2c848d-1ae7-11f0-80dc-00155d235309",
        "direction_name": "Логистика",
        "include_children": False,
        "reason": "ЦО по Логистике (док №2: не свод, просто привязка к направлению)",
    },
    "Экспедирование": {
        "struct_uuid": "9d2c8f03-1ae7-11f0-80dc-00155d235309",
        "struct_name": "Экспедирование",
        "reason": "Точное соответствие Экспедирование",
    },
    "MAN №3 white KA7837HO": {
        "struct_uuid": "9de7241f-1ae7-11f0-80dc-00155d235309",
        "struct_name": "MAN №3 white KA7837HO",
        "reason": "Каждая MAN-машина как отдельное подразделение в 1С",
    },
    "MAN №4 white KA3597HO": {
        "struct_uuid": "9de723d6-1ae7-11f0-80dc-00155d235309",
        "struct_name": "MAN №4 white KA3597HO",
        "reason": "Каждая MAN-машина как отдельное подразделение в 1С",
    },
    "MAN №5 Тент МД": {
        "struct_uuid": "9d2c8b61-1ae7-11f0-80dc-00155d235309",
        "struct_name": "MAN №5 KA5697НВ",
        "reason": "В 1С машина обозначена номером госзнака (KA5697НВ), в листе Excel — описанием (Тент МД)",
    },
}

SUMMARY_SHEET_EXACT = {
    "PL Строительство Свод", "PL_Свод", "PL_ЦО", "ЦО_Підсумовування",
}

TOTAL_ROW_KEYWORDS = [
    "итого", "валовая прибыль",
    "маржинальный доход", "операционная прибыль",
    "прибыль после вычета", "прибыль до вычета",
    "чистый доход", "рентабельность",
    "прибыль в распоряжении компании",
    ", грн", ", %", ", тис.грн",
]

# Явные строки-комментарии финансиста (подгруппы/пояснения), которые иногда попадают
# в столбец B без числовых маркеров и не отлавливаются TOTAL_ROW_KEYWORDS / regex.
# При совпадении (по точному имени, casefold) строка классифицируется как summary и
# в каталог А_Статьи_PL не попадает.
EXTRA_COMMENT_NAMES = {
    "Затраты ЦО",
    # Комментарий финансиста, попавший в столбец B как "статья" в 2025 (Май-Август).
    # Содержит арифметику и опечатку "рахсод" — точно не статья каталога.
    "Доход 3,7+0,7-рахсод  8,8-4,5",
}

# Алиасы названий PL-статей в Excel → каноничное имя в Catalog.А_Статьи_PL.
# Используется при сборке 02_unique_articles, чтобы вариант "Арена сторонней техники"
# не создавался как новая статья, а сворачивался в существующую "Аренда сторонней техники".
ARTICLE_NAME_ALIASES = {
    "Арена сторонней техники": "Аренда сторонней техники",
    # 2024-варианты для существующей "Штраф/ мат помощь":
    "Штраф, Материальная помощь": "Штраф/ мат помощь",
    "Штраф/ мат.помощ": "Штраф/ мат помощь",
    # 2026-05-21: финансист пометил на удаление дубль "Штраф" (000000060),
    # все варианты сводим в единую статью "Штраф/ мат помощь" (000000059).
    "Штраф": "Штраф/ мат помощь",
}

# Deny-list имён PL-статей, которые НЕЛЬЗЯ создавать в Catalog.А_Статьи_PL при заливке.
# Используется как защита от случайного восстановления статьи, которую финансист удалил.
# Проверяется в 11_upload_articles.py перед mgr.СоздатьЭлемент().
#
# Если в Excel встречается такое имя — оно ДОЛЖНО быть покрыто ARTICLE_NAME_ALIASES
# (сворачиваться в каноничную статью). Если алиаса нет, а имя в deny-list — скрипт
# пропустит создание со SKIP-BLOCKED.
#
# Сравнение по точному имени (case-insensitive, с обрезкой пробелов).
BLOCKED_ARTICLE_NAMES = {
    "Штраф",  # 2026-05-21 удалена; вместо неё → "Штраф/ мат помощь"
}

PL_GROUP_NAMES = [
    "Операционный доход",
    "Себестоимость проданой продукции (переменные производственные затраты)",
    "Себестоимость проданной продукции (переменные производственные затраты)",
    "Общепроизводственные затраты",
    "Административные затраты",
    "Маркетинговые затраты",
    "Дополнительные расходы",
    "Дополнительные  расходы",
    "Налоги и сборы",
    "Финансовая деятельность",
]

FUZZY_AUTO_THRESHOLD = 85
FUZZY_CANDIDATE_THRESHOLD = 70

# Ручные override для связи PL-имя Excel → UUID ДДС (Справочник.СтатьиДвиженияДенежныхСредств).
# Применяется в 04_match_articles.py ПОСЛЕ fuzzy-алгоритма.
# Добавляем пары, когда fuzzy не даёт exact/auto-совпадение из-за разных языков/терминологии.
# Источник связей — mapping_master.csv (анализ реальных данных АДР за декабрь 2025).
MANUAL_DDS_OVERRIDES = {
    # PL (Excel имя): (DDS UUID, DDS имя)
    "Логистические затраты": (
        "bcd804e5-c7e0-11e8-a200-000c299fb278",
        "Транспортные расходы",
    ),
    "Телекомуникационные услуги": (
        "6f023e7b-cbbb-11e8-a200-000c299fb278",
        "Интернет",
    ),
    "Юридические услуги": (
        "be5cc1e2-c8d0-11e8-a200-000c299fb278",
        "Прочие юр услуги",
    ),
    "ИТ расходы": (
        "14d1494e-cbbe-11e8-a200-000c299fb278",
        "Обслуживание офисной техники",
    ),
    "Услуги банков": (
        "61b6b4bf-c7e0-11e8-a200-000c299fb278",
        "Комиссия банков и др.платежи по РКО",
    ),
}

# Ручные override для связи PL-имя Excel → UUID СтатьяДоходов (ПВХ).
# Применяется в 04b_match_statya_dohodov.py ПОСЛЕ fuzzy-алгоритма.
# Источник — реальные связи в Catalog.А_Статьи_PL (поле СтатьяДоходов).
MANUAL_STATYA_DOHODOV_OVERRIDES = {
    # "Выручка от продаж..." сматчивается fuzzy, но для надёжности фиксируем вручную.
    "Выручка от продаж (товаров, работ, услуг)": (
        "77981290-6ea2-11f0-a2de-bccbaebe2890",
        "Выручка от продаж",
    ),
    "Доход от прочей операционной деятельности": (
        "77981293-6ea2-11f0-a2de-bccbaebe2890",
        "Прочие доходи",
    ),
    "Финансовый доход": (
        "0caba0bd-01f3-11f1-a2e2-c54425f51b91",
        "Прочие доходы от Фин деятельности",
    ),
}

PL_GROUPS = []

# Соответствие: группа PL → тип статьи (значение перечисления А_ТипСтатьиPL).
# Ключи — нормализованные (casefold + схлопнутые пробелы) фрагменты названия группы.
# Используется в excel_parser для заполнения поля type_statii у каждой детали.
GROUP_TO_TYPE = {
    "операционный доход": "Доход",
    "финансовая деятельность": "Расход",  # спец-правило: "Финансовый доход" перекрывается по имени статьи
}

# Имя статьи (нормализованное фрагмент) → тип. Срабатывает в приоритете над GROUP_TO_TYPE.
ARTICLE_NAME_TYPE_OVERRIDES = {
    "финансовый доход": "Доход",
}

# Итоговые показатели, которые из Excel-детализации НЕ берутся в ТЧ ДанныеОтчета,
# а уходят в ТЧ ИтогиОбщие. Ключ — подстрока имени статьи; значение — ВидПоказателя.
GENERAL_TOTAL_PATTERNS = {
    "маржинальный доход, грн": "Сумма",
    "маржинальный доход, %": "Процент",
    "операционная прибыль, грн": "Сумма",
    "операционная прибыль, %": "Процент",
    "прибыль после вычета финансовых затрат": "Сумма",
    "прибыль до вычета налогов": "Сумма",
    "чистый доход": "Сумма",
    "прибыль в распоряжении компании": "Сумма",
    "рентабельность продукции, %": "Процент",
    "рентабельность, %": "Процент",
}

# Маркеры КОНЦА тела PL-отчёта в листе Excel. Всё, что расположено НИЖЕ этих итоговых
# строк (после «прибыли в распоряжении» / «рентабельности»), — служебные заметки
# финансиста (договоры подряда, объекты, расчёты), а НЕ статьи PL. excel_parser
# пропускает detail-строки после первого такого маркера (структурный фильтр фейков).
# Подстрочное совпадение (casefold). Реальные промежуточные итоги (валовая прибыль,
# маржинальный доход, операционная прибыль) сюда НЕ входят — они остаются в ИтогиОбщие.
REPORT_END_MARKERS = [
    "прибыль в распоряжении компании",
    "рентабельность",
]


def determine_type(article_name: str, group_name: str) -> str:
    """Вернуть значение перечисления А_ТипСтатьиPL для детальной строки PL."""
    def _norm(s):
        return " ".join(str(s or "").split()).casefold()

    an = _norm(article_name)
    for frag, tp in ARTICLE_NAME_TYPE_OVERRIDES.items():
        if frag in an:
            return tp
    gn = _norm(group_name)
    for frag, tp in GROUP_TO_TYPE.items():
        if frag in gn:
            return tp
    # По умолчанию — Расход (подавляющая часть статей PL — расходы)
    return "Расход"


def classify_summary(article_name: str):
    """Если строка — общий итог из низа листа Excel, вернуть ВидПоказателя; иначе None."""
    def _norm(s):
        return " ".join(str(s or "").split()).casefold()
    an = _norm(article_name)
    for frag, kind in GENERAL_TOTAL_PATTERNS.items():
        if frag in an:
            return kind
    return None


PL_HEADER_ROW = 4
PL_DATA_START_ROW = 5

COL_ARTICLE = "B"
COL_F1 = "C"
COL_F2 = "D"
COL_TOTAL = "E"
COL_PERCENT = "F"
COL_COMMENT = "G"
