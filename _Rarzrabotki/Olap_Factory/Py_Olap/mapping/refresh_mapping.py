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
    or 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
)

# Whitelist of objects we actually need for ETL. Generated from olap_data_sources_erp.md.
WHITELIST = [
    # --- Целевой регистр свёртки и его регистратор ---
    "РегистрСведений.А_ПланФактПроизводство_Свод",
    "Документ.А_Отчет_ПланФактныйПроизводство",
    # --- Перечисление-дискриминатор строк (FROZEN_ENUMS) ---
    "Перечисление.А_ТипыСтрокПланФактПроизводство",
    # --- Справочники под Dim-таблицы (FK факта) ---
    "Справочник.Организации",
    "Справочник.СтруктураПредприятия",
    "Справочник.А_ЭтапыРабот",
    "Справочник.Номенклатура",
    "Справочник.А_ОбщиеНазванияНоменклатуры",
    "Справочник.УпаковкиЕдиницыИзмерения",
    "Справочник.ФизическиеЛица",
    "Справочник.А_СтруктураСебестоимости",
    # --- Документы для Dim_Documents: расшифровка каждой цифры до первички ---
    "Документ.ПриобретениеТоваровУслуг",
    "Документ.А_ТабельУчетаРабочегоВремени",
    "Документ.А_ВыполнениеРаботМодулей",
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
