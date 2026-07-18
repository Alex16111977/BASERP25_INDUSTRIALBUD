"""Read all enum values (UUID -> Name) for whitelisted enums via COM.

1С stores enum values' names only in metadata, not in SQL backend tables.
This script fetches names through V83.COMConnector and saves them to
mapping/enum_values.json for use by enum_resolver dynamic mode.

Run after mapping/refresh_mapping.py when an enum's WHITELIST changes.
"""
import json
import sys
import io
from pathlib import Path

import win32com.client

if sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

# Enums where we need full UUID -> Name mapping (large or dynamic content)
ENUMS_TO_DUMP = [
    "ХозяйственныеОперации",
]

OUT_PATH = Path(__file__).parent / "enum_values.json"


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN)
    out = {}
    for enum_name in ENUMS_TO_DUMP:
        meta = erp.Метаданные.Перечисления.Найти(enum_name)
        if meta is None:
            print(f"WARN: enum {enum_name} not found")
            continue
        manager = getattr(erp.Перечисления, enum_name)
        values = {}
        # Iterate via SELECT — efficient for large enums
        q = erp.NewObject("Запрос")
        q.Text = f"ВЫБРАТЬ Ссылка, Порядок ИЗ Перечисление.{enum_name}"
        rows = q.Execute().Выгрузить()
        n = rows.Количество()
        for i in range(n):
            row = rows.Получить(i)
            ref = row.Ссылка
            uid = erp.String(ref.УникальныйИдентификатор()).replace("-", "").lower()
            # Get value name via metadata
            name = erp.String(ref.Метаданные().Имя)  # this is enum metadata name, not value
            # Actually need value's own name — use ПолучитьИмяПредопределенногоЗначения
            try:
                name = erp.ПолучитьИмяПредопределенногоЗначения(ref)
                # Strip "Перечисление.ХозяйственныеОперации." prefix
                name = name.split(".")[-1] if name else ""
            except Exception:
                # Fallback to Synonym
                name = erp.String(ref) or f"Op_{int(row.Порядок)}"
            values[uid] = name
        out[f"Перечисление.{enum_name}"] = values
        print(f"{enum_name}: loaded {len(values)} values")

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"Saved {OUT_PATH}")


if __name__ == "__main__":
    main()
