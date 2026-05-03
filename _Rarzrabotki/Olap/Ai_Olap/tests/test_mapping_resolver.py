"""mapping/baserp_storage.json must resolve key 1С objects."""
from ai_olap.utils.mapping_resolver import list_objects, resolve


def test_organizations_maps_to_reference_table():
    table, fields = resolve("Справочник.Организации")
    assert table.startswith("_Reference"), f"unexpected table {table}"
    assert "Ссылка" in fields
    assert fields["Ссылка"] == "_IDRRef"


def test_pl_swod_register_maps_to_inforg():
    table, fields = resolve("РегистрСведений.А_ОтчетPL_Свод")
    assert table.startswith("_InfoRg"), f"unexpected table {table}"
    # Глобино-2 acceptance gate hangs on this column being mapped.
    assert "СуммаЕРПГрн" in fields
    assert fields["СуммаЕРПГрн"].startswith("_Fld")


def test_required_objects_present():
    available = set(list_objects())
    required = {
        "Справочник.Организации",
        "Справочник.СтруктураПредприятия",
        "РегистрСведений.А_ОтчетPL_Свод",
        "РегистрСведений.А_ОтчетDDS_Свод",
        "Документ.А_ФинРез_PL",
        "Перечисление.А_ИсточникPL",
        "Перечисление.А_ИсточникDDS",
    }
    missing = required - available
    assert not missing, f"mapping JSON missing: {missing}"
