"""Refresh mapping/baserp_storage.json from a live BaseERP via COM.

Calls platform method `GetDBStorageStructureInfo` (rus: ПолучитьСтруктуруХраненияБазыДанных)
and writes the JSON used by ai_olap.utils.mapping_resolver.

Run when the configuration changes (table IDs are platform-generated).

    python mapping/refresh_mapping.py

Requires: pywin32 (already in requirements.txt).
"""
from __future__ import annotations

import datetime as dt
import json
import os
import sys
from pathlib import Path

import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

CONN = (
    os.environ.get("BASERP_COM_CONN")
    or 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
)

# Whitelist of objects we actually need for ETL. Generated from olap_data_sources_erp.md.
WHITELIST = [
    # Dim catalogs / chart-of-types
    "Справочник.Организации",
    "Справочник.СтруктураПредприятия",
    "Справочник.НаправленияДеятельности",
    "Справочник.Контрагенты",
    "Справочник.Партнеры",
    "Справочник.ДоговорыКонтрагентов",
    "Справочник.Номенклатура",
    "Справочник.Склады",  # Dim_Warehouses (Себестоимость → Fact_Balance.Warehouse_ID)
    "Справочник.ГруппыФинансовогоУчетаНоменклатуры",
    "Справочник.ФизическиеЛица",
    "Справочник.Пользователи",
    "Справочник.БанковскиеСчета",
    "Справочник.БанковскиеСчетаОрганизаций",
    # Cash account composite components for Dim_DenezhnyeSredstva (Stage v3 2026-05-08)
    "Справочник.Кассы",
    "Справочник.КассыККМ",
    "Справочник.ДенежныеДокументы",
    "Справочник.ДоговорыКредитовИДепозитов",
    "Справочник.ЭквайринговыеТерминалы",
    "Справочник.Валюты",
    "Справочник.СтатьиДвиженияДенежныхСредств",
    "ПланВидовХарактеристик.СтатьиРасходов",
    "ПланВидовХарактеристик.СтатьиДоходов",
    "Справочник.А_Статьи_PL",
    "Справочник.А_ГруппаСтатей_PL",
    # Enums (frozen identifiers + dynamic for ХозяйственныеОперации)
    "Перечисление.А_ИсточникPL",
    "Перечисление.А_ИсточникDDS",
    "Перечисление.А_РазделыCFS",
    "Перечисление.ТипыДвиженияДенежныхСредств",
    "Перечисление.ТипыДенежныхСредств",
    "Перечисление.ХозяйственныеОперации",
    # 2 new info registers + 1 OLAP drill-down register (Stage v2 2026-05-05)
    "РегистрСведений.А_ОтчетPL_Свод",
    "РегистрСведений.А_ОтчетDDS_Свод",
    "РегистрСведений.А_ДокументРасшифровка",
    # 2 new documents
    "Документ.А_ФинРез_PL",
    "Документ.А_ФинРез_DDS",
    # Balance Stage (2026-05-16): А_ОтчетБаланс_Свод -> Fact_Balance
    "РегистрСведений.А_ОтчетБаланс_Свод",
    "Документ.А_ФинРез_Баланс",
    "ПланВидовХарактеристик.СтатьиАктивовПассивов",
    "Справочник.ОбъектыРасчетов",  # Dim_ObjektyRaschetov (Свод_РасчетыСПартнерами → Fact_Balance.SettlementObj_ID)
    "Справочник.А_ФинАгенты",  # Dim_FinAgents (ДоговорыКонтрагентов.А_ФинАгент)
    "Справочник.ИдентификаторыОбъектовМетаданных",  # тип composite Объект (ОбъектыРасчетов.ТипСсылки)
    "Перечисление.А_ИсточникБаланса",
    "Перечисление.ВидыСтатейУправленческогоБаланса",
    # А_ОтчетБаланс_Свод.Source factually uses the PLATFORM enum (not custom)
    "Перечисление.ИсточникиУправленческогоБаланса",
    # Balance Прямой (2026-05-18): А_ОтчетБаланс_Свод.ТипНалога → Fact_Balance.TaxType
    "Перечисление.ТипыНалогов",
    "Перечисление.ТипыДоговоров",            # Dim_Contracts.TipDogovora
    "Перечисление.ТипыРасчетовСПартнерами",  # Dim_ObjektyRaschetov.TipRaschetov
    "Перечисление.ТипыОбъектовРасчетов",     # Dim_ObjektyRaschetov.TipObjektaRaschetov
    # source registers needed for COM fallback
    "РегистрНакопления.ДенежныеСредстваБезналичные",
    "РегистрНакопления.ДенежныеСредстваНаличные",
    # subdivisions (for departments hierarchy)
    "Справочник.Подразделения",
]


def serialize_struct(struct) -> dict:
    """Convert a COM array of ОписаниеТаблиц items into a JSON-serializable dict."""
    objects: dict[str, dict] = {}
    count = struct.Количество()
    for i in range(count):
        item = struct.Получить(i)
        meta_full = str(item.Метаданные) if hasattr(item, "Метаданные") else ""
        if not meta_full:
            continue
        if meta_full not in objects:
            objects[meta_full] = {"primary_table": None, "tables": [], "fields": {}}

        table_name = str(item.ИмяТаблицыХранения)
        purpose = str(item.Назначение) if hasattr(item, "Назначение") else "Main"

        fields_dump: dict[str, str] = {}
        if hasattr(item, "Поля"):
            fields_arr = item.Поля
            for j in range(fields_arr.Количество()):
                fld = fields_arr.Получить(j)
                fname = str(fld.ИмяПоля)
                fstore = str(fld.ИмяПоляХранения)
                fields_dump[fname] = fstore

        objects[meta_full]["tables"].append(
            {"name": table_name, "purpose": purpose, "fields": fields_dump}
        )
        if purpose in ("Основная", "Main") or objects[meta_full]["primary_table"] is None:
            objects[meta_full]["primary_table"] = table_name
            objects[meta_full]["fields"] = fields_dump

    return objects


def main() -> int:
    print(f"Connecting via COM to {CONN[:60]}…")
    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN)
    print("Connected.")

    print(f"Calling GetDBStorageStructureInfo for {len(WHITELIST)} whitelisted objects…")

    array_class = erp.NewObject("Массив")
    for full_name in WHITELIST:
        try:
            mo = erp.Метаданные.НайтиПоПолномуИмени(full_name)
            if mo is None:
                print(f"  WARN: not found {full_name}")
                continue
            array_class.Добавить(mo)
        except Exception as exc:
            print(f"  WARN: lookup failed {full_name}: {exc!r}")

    # GetDBStorageStructureInfo / ПолучитьСтруктуруХраненияБазыДанных
    struct = erp.ПолучитьСтруктуруХраненияБазыДанных(array_class, True)
    print(f"Struct items: {struct.Количество()}")

    objects = serialize_struct(struct)
    print(f"Serialized {len(objects)} top-level metadata objects.")

    # Augment with engine version
    version = ""
    try:
        version = str(erp.СистемнаяИнформация().ВерсияПриложения)
    except Exception:
        pass

    out = {
        "version": version,
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "source_db": "BaseERP",
        "object_count": len(objects),
        "objects": objects,
    }

    out_path = Path(__file__).parent / "baserp_storage.json"
    out_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {out_path}  ({out_path.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
