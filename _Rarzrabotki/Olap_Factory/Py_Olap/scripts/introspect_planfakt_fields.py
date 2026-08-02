# -*- coding: utf-8 -*-
"""
Интроспекция физических имён (_InfoRgNNNNN / _FldNNN) для контура План-факт производства.
Аналог scripts/introspect_balance_fields.py из Ai_Olap.

Печатает готовый SELECT-список для raw_sql пайплайна — чтобы номера полей не набивались руками.
После КАЖДОЙ перегенерации mapping/baserp_storage.json прогонять заново: платформенные _FldNNN
зависят от порядка реквизитов и могут сдвинуться.

Запуск:  .venv/Scripts/python.exe scripts/introspect_planfakt_fields.py
"""
import json, sys, os
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
ROOT = Path(__file__).resolve().parent.parent
MAPPING = ROOT / "mapping" / "baserp_storage.json"

ОБЪЕКТЫ = [
    "РегистрСведений.А_ПланФактПроизводство_Свод",
    "Документ.А_Отчет_ПланФактныйПроизводство",
    "Перечисление.А_ТипыСтрокПланФактПроизводство",
    "Справочник.Организации",
    "Справочник.СтруктураПредприятия",
    "Справочник.А_ЭтапыРабот",
    "Справочник.Номенклатура",
    "Справочник.А_ОбщиеНазванияНоменклатуры",
    "Справочник.УпаковкиЕдиницыИзмерения",
]


def найти(данные, полное_имя):
    if isinstance(данные, dict):
        if полное_имя in данные:
            return данные[полное_имя]
        for v in данные.values():
            r = найти(v, полное_имя)
            if r is not None:
                return r
    elif isinstance(данные, list):
        for v in данные:
            r = найти(v, полное_имя)
            if r is not None:
                return r
    return None


def main():
    if not MAPPING.exists():
        print(f"НЕТ {MAPPING} — сначала прогони mapping/refresh_mapping.py")
        return 1
    данные = json.loads(MAPPING.read_text(encoding="utf-8"))

    for имя in ОБЪЕКТЫ:
        узел = найти(данные, имя)
        print("=" * 100)
        print(имя)
        if узел is None:
            print("   НЕ НАЙДЕН в маппинге — добавь в WHITELIST refresh_mapping.py и перегенерируй")
            continue
        печать_узла(узел)
    return 0


def печать_узла(узел, отступ="   "):
    if isinstance(узел, dict):
        табл = узел.get("tableName") or узел.get("TableName") or узел.get("Имя")
        if табл:
            print(f"{отступ}таблица: {табл}")
        поля = узел.get("fields") or узел.get("Fields") or узел.get("Поля")
        if isinstance(поля, dict):
            for k, v in поля.items():
                кол = v.get("columnName") if isinstance(v, dict) else v
                print(f"{отступ}  {k:34} -> {кол}")
        elif isinstance(поля, list):
            for f in поля:
                if isinstance(f, dict):
                    print(f"{отступ}  {str(f.get('name') or f.get('Имя')):34} -> "
                          f"{f.get('columnName') or f.get('ИмяПоля')}")
        for k, v in узел.items():
            if k in ("fields", "Fields", "Поля", "tableName", "TableName", "Имя"):
                continue
            if isinstance(v, (dict, list)):
                print(f"{отступ}{k}:")
                печать_узла(v, отступ + "   ")
    elif isinstance(узел, list):
        for v in узел:
            печать_узла(v, отступ)


if __name__ == "__main__":
    raise SystemExit(main())
