"""Правка реквизита СтатьяДвиженияДенежныхСредств (шапка + ТЧ.Статьи стр.1)
для 5 PL-статей, у которых была обобщённая ДДС (ЗП / Персонал).

После правки JOIN в отчёте А_ОтчетPL (ДДС) заработает и даст СуммаЕРП для ЗП.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.com_connect import connect_erp, is_manually_locked

# (PL_UUID, PL_name, DDS_UUID, DDS_name)
FIXES = [
    ("78e2eb1c-39d1-11f1-a2e8-ae822d4b0de9", "ЗП производственного персонала",
     "7bda3c70-c7e6-11e8-a200-000c299fb278", "Зарплата производственного персонала"),
    ("78e2eb22-39d1-11f1-a2e8-ae822d4b0de9", "ЗП подсобников",
     "4b202afd-cbbf-11e8-a200-000c299fb278", "Зарплата подсобников"),
    ("78e2eb3f-39d1-11f1-a2e8-ae822d4b0de9", "ЗП управленческого персонала",
     "56fdb2ca-c894-11e8-a200-000c299fb278", "Зарплата управленческого персонала"),
    ("78e2eb42-39d1-11f1-a2e8-ae822d4b0de9", "ЗП РМ",
     "5b124159-2308-11e9-a203-000c299fb278", "Зарплата PM"),
    ("78e2eb41-39d1-11f1-a2e8-ae822d4b0de9", "Удержание с зарплаты управленческого персонала",
     "7a52de2f-2df8-11e9-a203-000c299fb278", "Удержание с админ зарплату"),
]


def main():
    conn = connect_erp()
    pl_mgr = conn.Справочники.А_Статьи_PL
    dds_mgr = conn.Справочники.СтатьиДвиженияДенежныхСредств

    for pl_uuid, pl_name, dds_uuid, dds_name in FIXES:
        pl_ref = pl_mgr.ПолучитьСсылку(conn.NewObject("УникальныйИдентификатор", pl_uuid))
        dds_ref = dds_mgr.ПолучитьСсылку(conn.NewObject("УникальныйИдентификатор", dds_uuid))
        obj = pl_ref.ПолучитьОбъект()
        if obj is None:
            print(f"  SKIP: {pl_name} — объект не найден")
            continue
        if is_manually_locked(obj):
            print(f"  LOCK {pl_name!r} — статья защищена (А_РучнаяКорректировка=True)")
            continue

        old_dds_header = str(obj.СтатьяДвиженияДенежныхСредств) if obj.СтатьяДвиженияДенежныхСредств else "(пусто)"
        obj.СтатьяДвиженияДенежныхСредств = dds_ref

        # Обновить первую строку ТЧ Статьи (если её нет — добавить)
        tch = obj.Статьи
        if tch.Количество() == 0:
            row = tch.Добавить()
        else:
            row = tch.Получить(0)
        old_dds_tch = str(row.СтатьяДвиженияДенежныхСредств) if row.СтатьяДвиженияДенежныхСредств else "(пусто)"
        row.СтатьяДвиженияДенежныхСредств = dds_ref

        obj.Записать()
        print(f"  UPD {pl_name!r:55}")
        print(f"      Шапка: {old_dds_header!r:40} -> {dds_name!r}")
        print(f"      ТЧ[0]: {old_dds_tch!r:40} -> {dds_name!r}")

    print("\nГотово. Запустить верификацию:")
    print("  python scripts/_test_otchet_pl_astarta.py")


if __name__ == "__main__":
    main()
