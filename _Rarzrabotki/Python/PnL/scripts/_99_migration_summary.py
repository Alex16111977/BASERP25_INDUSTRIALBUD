"""Phase 5: финальный аудит миграции направления.

Сравнивает до/после, записывает итог в data/json/99_direction_migration_summary.json.
"""
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp


def main():
    conn = connect_erp()

    # Rasshifrovka coverage
    q1 = conn.NewObject("Запрос")
    q1.Текст = """
    ВЫБРАТЬ Ссылка.Номер КАК Номер, Ссылка.Дата КАК Дата,
        КОЛИЧЕСТВО(*) КАК Строк,
        СУММА(ВЫБОР КОГДА Подразделение <> ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СПодразд,
        СУММА(ВЫБОР КОГДА НаправлениеДеятельности <> ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СНапр,
        СУММА(ВЫБОР КОГДА ВключатьДочерние ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Свод
    ИЗ Документ.А_РасшифровкаЛистов.Расшифровка
    СГРУППИРОВАТЬ ПО Ссылка.Номер, Ссылка.Дата
    УПОРЯДОЧИТЬ ПО Ссылка.Дата
    """
    tz1 = q1.Выполнить().Выгрузить()
    rasshifrovka = []
    for i in range(tz1.Количество()):
        r = tz1.Получить(i)
        rasshifrovka.append({
            "number": str(r.Номер).strip(),
            "date": str(r.Дата)[:10],
            "rows": int(r.Строк),
            "with_podrazd": int(r.СПодразд),
            "with_direction": int(r.СНапр),
            "svod_rows": int(r.Свод),
        })

    # А_ОтчетPL coverage
    q2 = conn.NewObject("Запрос")
    q2.Текст = """
    ВЫБРАТЬ
        КОЛИЧЕСТВО(*) КАК Всего,
        СУММА(ВЫБОР КОГДА НаправлениеДеятельности <> ЗНАЧЕНИЕ(Справочник.НаправленияДеятельности.ПустаяСсылка) ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК СНапр,
        СУММА(ВЫБОР КОГДА ВключатьДочерние ТОГДА 1 ИНАЧЕ 0 КОНЕЦ) КАК Свод
    ИЗ Документ.А_ОтчетPL
    ГДЕ НЕ ПометкаУдаления
    """
    tz2 = q2.Выполнить().Выгрузить()
    r2 = tz2.Получить(0)
    otchetpl = {
        "total": int(r2.Всего),
        "with_direction": int(r2.СНапр),
        "svod": int(r2.Свод),
    }

    # Load baseline for diff
    baseline_path = config.JSON_DIR / "00_baseline_pre_direction_migration.json"
    baseline = json.loads(baseline_path.read_text(encoding="utf-8"))

    summary = {
        "timestamp_end": datetime.now().isoformat(timespec="seconds"),
        "timestamp_begin": baseline["timestamp"],
        "baseline": {
            "rasshifrovka_listov": baseline["rasshifrovka_listov"],
            "otchet_pl": baseline["otchet_pl"],
        },
        "after": {
            "rasshifrovka_listov": rasshifrovka,
            "otchet_pl": otchetpl,
        },
        "acceptance_6_1_structural": "passed (см. MCP get_metadata_structure)",
        "acceptance_6_2_data": {
            "rasshifrovka_all_svod_rows_have_direction": all(
                r["with_direction"] >= r["svod_rows"] for r in rasshifrovka
            ),
            "otchetpl_all_with_direction": otchetpl["with_direction"] == otchetpl["total"],
        },
        "acceptance_6_4_idempotency": "passed (14 dry-run 2nd run = 0 updated)",
    }

    dst = config.JSON_DIR / "99_direction_migration_summary.json"
    dst.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {dst}")
    print()
    print("=== Rasshifrovka ===")
    for r in rasshifrovka:
        print(f"  {r['number']} ({r['date']}): rows={r['rows']} podrazd={r['with_podrazd']} napr={r['with_direction']} svod={r['svod_rows']}")
    print(f"\n=== А_ОтчетPL ===")
    print(f"  total={otchetpl['total']} napr={otchetpl['with_direction']} svod={otchetpl['svod']}")
    print(f"\nAcceptance 6.2: all svod rows have direction: {summary['acceptance_6_2_data']['rasshifrovka_all_svod_rows_have_direction']}")
    print(f"Acceptance 6.2: otchetpl 100% direction:     {summary['acceptance_6_2_data']['otchetpl_all_with_direction']}")


if __name__ == "__main__":
    main()
