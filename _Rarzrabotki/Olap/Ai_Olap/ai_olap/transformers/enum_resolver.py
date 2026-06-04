"""Resolve 1С Enum refs to frozen string identifiers.

The Power BI model relies on these strings literally — never rename them.

For each known enum:
1) build a one-shot {hex_uuid: frozen_string} map by reading _IDRRef + _EnumOrder
   from the SQL backend (table from mapping_resolver) and zipping with the
   frozen list defined here;
2) cache the map for the process lifetime;
3) replace UUIDs in input rows.

If a row has a UUID not present in the map, the value becomes None (broken ref)
and a warning is logged.

Dynamic enums (DYNAMIC_ENUMS): for enums that are too large to maintain frozen
(>100 values, e.g. ХозяйственныеОперации has ~600 values), build the map by reading
_IDRRef + _Description directly from the SQL backend. No frozen list required.
"""
from __future__ import annotations

import functools

import structlog

from ..core.connections import get_baserp_sql
from ..core.exceptions import TransformError
from ..utils.mapping_resolver import resolve

log = structlog.get_logger().bind(component="enum_resolver")


# Frozen enum value lists — DO NOT REORDER; Power BI DAX references these literals.
# Order MUST match _EnumOrder values in 1C metadata.
FROZEN_ENUMS: dict[str, list[str]] = {
    "Перечисление.А_ИсточникPL": [
        "PL_Excel",
        "PL_ЕРП",
    ],
    # 2026-05-21: ТипСтатьи у Справочник.А_Статьи_PL → Dim_PL_Articles.Type_Statya.
    # Семантика знака суммы в А_ФинРез_PL: Доход → +Сумма, Расход → -Сумма.
    # Order MUST match Metadata.Enums.А_ТипСтатьиPL._EnumOrder (live read 2026-05-21).
    "Перечисление.А_ТипСтатьиPL": [
        "Доход",              # order 0
        "Расход",             # order 1
        "ОперационныйИтог",   # order 2 (для итоговых строк отчёта)
        "Информационный",     # order 3 (для справочных статей)
    ],
    # Stage v3 (2026-05-08): rebuild А_ИсточникDDS to 4 values matching new register
    "Перечисление.А_ИсточникDDS": [
        "ЕРП",          # order 0
        "Казна",        # order 1
        "План",         # order 2
        "ПланОбъекта",  # order 3
    ],
    "Перечисление.А_РазделыCFS": [
        "Operating",
        "Investing",
        "Financing",
        "Internal",
    ],
    "Перечисление.ТипыДвиженияДенежныхСредств": [
        "Поступление",  # order 0
        "Списание",     # order 1
    ],
    "Перечисление.ТипыДенежныхСредств": [
        "Наличные",                          # order 0
        "Безналичные",                       # order 1
        "ДенежныеСредстваУЭквайера",         # order 2
        "ДенежныеСредстваУПодотчетногоЛица", # order 3
        "Депозиты",                          # order 4
        "ДенежныеСредстваВПути",             # order 5
        "ДенежныеДокументы",                 # order 6
    ],
    # Balance Stage (2026-05-16): Fact_Balance Source. Order MUST match
    # Enums/А_ИсточникБаланса.xml EnumValue sequence (_EnumOrder 0..6).
    "Перечисление.А_ИсточникБаланса": [
        "ПрочиеАктивыПассивы",         # order 0
        "РасчетыСКлиентами",           # order 1
        "РасчетыСПоставщиками",        # order 2
        "СебестоимостьТоваров",        # order 3
        "ДенежныеСредстваБезналичные", # order 4
        "ДенежныеСредстваНаличные",    # order 5
        "ПрочиеРасходы",               # order 6
    ],
    # Dim_PAP_Articles.AktivPassiv (канон OD-9). ASCII identifiers — referenced
    # literally by PL.pbix DAX ([AktivPassiv]="Aktiv"|"Passiv"|"AktivPassiv")
    # and stored in varchar(15). Order MUST match
    # Enums/ВидыСтатейУправленческогоБаланса.xml (_EnumOrder 0..2).
    # МЕТАИМЕНА 1С (как Source/ТипыНалогов — конвенция пайплайна; FK
    # Fact_Balance.TipPokazatelya == Dim_TipPokazatelya.TipPokazatelya,
    # seed из 1С COM Имя). Регистр.ТипПоказателя = только Актив/Пассив
    # (формула УпрБаланс АктивПассив→Пассив). _EnumOrder 0..2.
    "Перечисление.ВидыСтатейУправленческогоБаланса": [
        "Актив",        # order 0
        "Пассив",       # order 1
        "АктивПассив",  # order 2 (в регистр не пишется — формула→Пассив)
    ],
    # Balance Source: register А_ОтчетБаланс_Свод.Source factually stores the
    # PLATFORM enum ИсточникиУправленческогоБаланса (NOT the custom
    # А_ИсточникБаланса from spec v3). 31 platform values; order MUST match
    # _EnumOrder 0..30 (read live from Metadata.Enums on 2026-05-17).
    "Перечисление.ИсточникиУправленческогоБаланса": [
        "АмортизацияНМА",                            # 0
        "АмортизацияОС",                             # 1
        "ДенежныеДокументы",                         # 2
        "ДенежныеСредстваБезналичные",               # 3
        "ДенежныеСредстваВКассахККМ",                # 4
        "ДенежныеСредстваВПути",                     # 5
        "ДенежныеСредстваНаличные",                  # 6
        "ДенежныеСредстваУПодотчетныхЛиц",           # 7
        "ПартииПрочихРасходов",                      # 8
        "ПереданнаяВозвратнаяТара",                  # 9
        "ПодарочныеСертификаты",                     # 10
        "ПринятаяВозвратнаяТара",                    # 11
        "ПрочиеАктивыПассивы",                       # 12
        "ПрочиеДоходы",                              # 13
        "ПрочиеРасходы",                             # 14
        "РасчетыПоФинансовымИнструментам",           # 15
        "РасчетыСКлиентами",                         # 16
        "РасчетыСКлиентамиПоДокументам",             # 17
        "РасчетыСПоставщиками",                      # 18
        "РасчетыСКлиентамиПоСрокам",                 # 19
        "РасчетыСПоставщикамиПоДокументам",          # 20
        "СебестоимостьТоваров",                      # 21
        "РасчетыСПоставщикамиПоСрокам",              # 22
        "РезервыПоСомнительнымДолгам",               # 23
        "СтоимостьНМА",                              # 24
        "СтоимостьОС",                               # 25
        "ТоварыКОформлениюОтчетовКомитенту",         # 26
        "ТрудозатратыНезавершенногоПроизводства",    # 27
        "ПрочиеРасходыНезавершенногоПроизводства",   # 28
        "УслугиКОформлениюОтчетовПринципалу",        # 29
        "РозничныеПродажиПодакцизныхТоваров",        # 30
    ],
    # Balance Прямой (2026-05-18): Fact_Balance.TaxType из ПАП.Аналитика
    # ВЫРАЗИТЬ(... КАК Перечисление.ТипыНалогов) для статьи «Налоги»
    # (Источник=ПустаяСсылка). Order MUST match Metadata.Enums.ТипыНалогов
    # ЗначенияПеречисления (_EnumOrder 0..13, read live 2026-05-18).
    "Перечисление.ТипыНалогов": [
        "НДС",                                       # 0
        "НДФЛ",                                      # 1
        "НДФЛДоходыКонтрагентов",                    # 2
        "НФДЛДивиденды",                             # 3
        "НФДЛДивидендыСотрудникам",                  # 4
        "НДФЛДоначисленныйПоРезультатамПроверки",    # 5
        "НДФЛПередачаЗадолженностиВНалоговыйОрган",  # 6
        "НДФЛПрочиеРасчетыСПерсоналом",              # 7
        "НачисленныйЕСВ",                            # 8
        "ВоенныйСбор",                               # 9
        "ЕдиныйНалог",                               # 10
        "НалогНаПрибыль",                            # 11
        "ДругиеНалоги",                              # 12
        "Акциз",                                     # 13
    ],
    # 2026-05-19: Dim_Contracts/Dim_ObjektyRaschetov реквизиты. Порядок ==
    # Metadata.Enums.<name>.ЗначенияПеречисления._EnumOrder (gen_frozen_enums_contracts.py).
    "Перечисление.ТипыДоговоров": [
        "СПокупателем",  # 0
        "СКомиссионером",  # 1
        "СПоставщиком",  # 2
        "СКомитентом",  # 3
        "Импорт",  # 4
        "СДавальцем",  # 5
        "СПереработчиком",  # 6
        "СПоклажедателем",  # 7
        "СХранителем",  # 8
        "СКомитентомНаЗакупку",  # 9
        "ИмпортКомиссия",  # 10
    ],
    "Перечисление.ТипыРасчетовСПартнерами": [
        "РасчетыСПоставщиком",  # 0
        "РасчетыСКлиентом",  # 1
        "РасчетыСКредитором",  # 2
        "РасчетыСДебитором",  # 3
        "РасчетыСЛизингодателем",  # 4
    ],
    "Перечисление.ТипыОбъектовРасчетов": [
        "Заказ",  # 0
        "Договор",  # 1
        "Накладная",  # 2
        "ПлатежВозврат",  # 3
    ],
}


# Dynamic enums — too large for frozen list. Resolver reads _Description directly.
# Use 1C metadata name (synonym) as the string identifier.
DYNAMIC_ENUMS: set[str] = {
    "Перечисление.ХозяйственныеОперации",
}


@functools.lru_cache(maxsize=16)
def _load_enum_map(meta_full: str) -> dict[str, str]:
    """Build {hex_uuid: frozen_string} for an enum metadata path.

    For frozen enums — zip _EnumOrder with FROZEN_ENUMS[meta_full].
    For dynamic enums — read _Description directly (no frozen list).
    """
    if meta_full in FROZEN_ENUMS:
        return _load_frozen_enum_map(meta_full)
    elif meta_full in DYNAMIC_ENUMS:
        return _load_dynamic_enum_map(meta_full)
    else:
        raise TransformError(f"No enum mapping defined for {meta_full}")


def _load_frozen_enum_map(meta_full: str) -> dict[str, str]:
    table, fields = resolve(meta_full)
    idref = fields["Ссылка"]
    order = fields["Порядок"]
    out: dict[str, str] = {}
    with get_baserp_sql() as c:
        cur = c.cursor()
        cur.execute(f"SELECT {idref}, {order} FROM {table}")
        for row in cur.fetchall():
            uid: bytes = row[0]
            n: int = int(row[1])
            try:
                value = FROZEN_ENUMS[meta_full][n]
            except IndexError as exc:
                raise TransformError(
                    f"{meta_full}: enum order {n} out of frozen list "
                    f"(size {len(FROZEN_ENUMS[meta_full])}). Configuration changed?"
                ) from exc
            out[uid.hex()] = value
    log.info("frozen enum loaded", enum=meta_full, count=len(out))
    return out


def _load_dynamic_enum_map(meta_full: str) -> dict[str, str]:
    """Read all UUID -> Description (Synonym/Name) for large enums.

    1C platform stores enum value name in `_Description` column of the _Enum table.
    """
    table, fields = resolve(meta_full)
    idref = fields["Ссылка"]
    out: dict[str, str] = {}
    with get_baserp_sql() as c:
        cur = c.cursor()
        # _Description holds the enum value Synonym (display name)
        cur.execute(f"SELECT {idref}, _Description FROM {table}")
        for row in cur.fetchall():
            uid: bytes = row[0]
            descr: str = (row[1] or "").strip()
            out[uid.hex()] = descr
    log.info("dynamic enum loaded", enum=meta_full, count=len(out))
    return out


def reload_cache() -> None:
    _load_enum_map.cache_clear()


def transform(rows: list[dict], *, column_to_enum: dict[str, str]) -> list[dict]:
    """Replace UUID hex strings (or bytes) with frozen strings.

    column_to_enum: {row_key: 1C_metadata_full_name}, e.g.
        {"Source": "Перечисление.А_ИсточникPL"}
    """
    maps = {col: _load_enum_map(meta) for col, meta in column_to_enum.items()}
    for row in rows:
        for col, m in maps.items():
            if col not in row:
                continue
            val = row[col]
            if val is None:
                # varbinary_to_uuid мапит пустую ссылку (16 нулевых байт) в
                # None ДО enum_resolver. Для enum-колонки это «пустая ссылка
                # перечисления» → 1С-имя "ПустаяСсылка" (НЕ NULL: Fact_Balance
                # .Source NOT NULL; согласовано с BSL Свод_*
                # ЗНАЧЕНИЕ(Перечисление...ПустаяСсылка) и Source-слайсером).
                row[col] = "ПустаяСсылка"
                continue
            if isinstance(val, (bytes, bytearray, memoryview)):
                val = bytes(val).hex()
            row[col] = m.get(val)
    return rows
