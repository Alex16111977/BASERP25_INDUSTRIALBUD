"""Этап 1.2 — выявляет непокрытые ДДС, дубли ДДС, PL без ДДС.

Читает:
- data/json/adr_dds_universe.json (от _collect_adr_dds.py)
- Справочник А_Статьи_PL + ТЧ Статьи (из 1С через COM)

Генерирует 3 CSV:
- data/reports/dds_uncovered.csv — ДДС из АДР без PL-покрытия
- data/reports/dds_duplicate.csv — ДДС в 2+ PL-статьях
- data/reports/pl_without_dds.csv — PL без шапочной ДДС или с пустой ТЧ
"""
import csv
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.com_connect import connect_erp, uuid_str


def collect_pl_catalog(conn):
    """Вытянуть все PL-статьи с шапкой + ТЧ Статьи."""
    # Шапка
    q_head = conn.NewObject("Запрос")
    q_head.Текст = """
    ВЫБРАТЬ
        PL.Ссылка КАК Ссылка,
        PL.Наименование КАК Имя,
        PL.Код КАК Код,
        PL.Группа.Наименование КАК Группа,
        PL.ТипСтатьи КАК ТипСтатьи,
        PL.СтатьяДвиженияДенежныхСредств КАК ДДСШапка,
        PL.А_СозданоАвтоматически КАК Авто
    ИЗ Справочник.А_Статьи_PL КАК PL
    ГДЕ НЕ PL.ПометкаУдаления И НЕ PL.ЭтоГруппа
    """
    tz_head = q_head.Выполнить().Выгрузить()
    pl_items = {}  # uuid → {...}
    for i in range(tz_head.Количество()):
        r = tz_head.Получить(i)
        pl_uuid = uuid_str(conn, r.Ссылка)
        шапка_dds_uuid = uuid_str(conn, r.ДДСШапка) if r.ДДСШапка and not r.ДДСШапка.Пустая() else None
        pl_items[pl_uuid] = {
            "uuid": pl_uuid,
            "name": str(r.Имя),
            "code": str(r.Код),
            "group": str(r.Группа) if r.Группа else "",
            "тип": str(r.ТипСтатьи) if r.ТипСтатьи else "",
            "шапка_dds": шапка_dds_uuid,
            "авто": bool(r.Авто),
            "тч_dds": [],  # список UUID ДДС из ТЧ
        }

    # ТЧ Статьи
    q_tch = conn.NewObject("Запрос")
    q_tch.Текст = """
    ВЫБРАТЬ
        ТЧ.Ссылка КАК Ссылка,
        ТЧ.НомерСтроки КАК НомерСтроки,
        ТЧ.СтатьяДвиженияДенежныхСредств КАК ДДС,
        ТЧ.СтатьяДвиженияДенежныхСредств.Наименование КАК ДДСИмя
    ИЗ Справочник.А_Статьи_PL.Статьи КАК ТЧ
    ГДЕ НЕ ТЧ.Ссылка.ПометкаУдаления И НЕ ТЧ.Ссылка.ЭтоГруппа
    УПОРЯДОЧИТЬ ПО ТЧ.Ссылка, ТЧ.НомерСтроки
    """
    tz_tch = q_tch.Выполнить().Выгрузить()
    for i in range(tz_tch.Количество()):
        r = tz_tch.Получить(i)
        pl_uuid = uuid_str(conn, r.Ссылка)
        dds_uuid = uuid_str(conn, r.ДДС) if r.ДДС and not r.ДДС.Пустая() else None
        if pl_uuid in pl_items and dds_uuid:
            pl_items[pl_uuid]["тч_dds"].append({"uuid": dds_uuid, "name": str(r.ДДСИмя or "")})

    return pl_items


def main():
    # 1. Загрузить adr_dds_universe.json
    universe = json.loads((config.JSON_DIR / "adr_dds_universe.json").read_text(encoding="utf-8"))

    # 2. Собрать PL-каталог
    conn = connect_erp()
    pl_items = collect_pl_catalog(conn)
    print(f"PL-статей всего: {len(pl_items)}")
    авто_ct = sum(1 for p in pl_items.values() if p["авто"])
    print(f"  из них авто-созданных: {авто_ct}")

    # 3. Построить индекс: dds_uuid → [список PL uuid где эта ДДС встречается в ТЧ]
    dds_to_pl = {}  # dds_uuid → [pl_uuid, ...]
    for pl_uuid, pl in pl_items.items():
        for ts in pl["тч_dds"]:
            dds_to_pl.setdefault(ts["uuid"], []).append(pl_uuid)

    print(f"ДДС использованных в ТЧ всего: {len(dds_to_pl)}")

    config.REPORTS_DIR.mkdir(parents=True, exist_ok=True)

    # 4. dds_uncovered.csv — ДДС из АДР, которых нет ни в одной PL.ТЧ
    uncovered_rows = []
    for ent in universe["расходы_ДДС"]:
        if ent["uuid"] == "__ПУСТО__":
            continue  # записи без ДДС в ПВХ — пропускаем, это отдельная тема
        if ent["uuid"] not in dds_to_pl:
            uncovered_rows.append({
                "ДДС_UUID": ent["uuid"],
                "ДДС_Код": ent["code"],
                "ДДС_Имя": ent["name"],
                "Итого_3мес_₴": ent["итого"],
                "Подразделений": len(ent["подразделений"]),
                "Примеры_подр": "; ".join(ent["подразделений"][:3]),
            })

    uncov_path = config.REPORTS_DIR / "dds_uncovered.csv"
    with uncov_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["ДДС_UUID", "ДДС_Код", "ДДС_Имя", "Итого_3мес_₴",
                                         "Подразделений", "Примеры_подр"], delimiter=";")
        w.writeheader()
        for r in sorted(uncovered_rows, key=lambda x: -x["Итого_3мес_₴"]):
            r["Итого_3мес_₴"] = f"{r['Итого_3мес_₴']:.2f}"
            w.writerow(r)

    # 5. dds_duplicate.csv — ДДС в 2+ PL
    dup_rows = []
    for dds_uuid, pl_list in dds_to_pl.items():
        if len(pl_list) > 1:
            pl_names = [pl_items[u]["name"] for u in pl_list]
            # найти имя ДДС
            dds_name = ""
            for pl in pl_items.values():
                for ts in pl["тч_dds"]:
                    if ts["uuid"] == dds_uuid:
                        dds_name = ts["name"]
                        break
                if dds_name:
                    break
            # найти сумму из universe
            sum_uni = 0
            for ent in universe["расходы_ДДС"]:
                if ent["uuid"] == dds_uuid:
                    sum_uni = ent["итого"]
                    break
            dup_rows.append({
                "ДДС_UUID": dds_uuid,
                "ДДС_Имя": dds_name,
                "Количество_PL": len(pl_list),
                "PL_статьи": " || ".join(pl_names),
                "Итого_3мес_₴": sum_uni,
            })

    dup_path = config.REPORTS_DIR / "dds_duplicate.csv"
    with dup_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["ДДС_UUID", "ДДС_Имя", "Количество_PL",
                                         "PL_статьи", "Итого_3мес_₴"], delimiter=";")
        w.writeheader()
        for r in sorted(dup_rows, key=lambda x: -x["Количество_PL"]):
            r["Итого_3мес_₴"] = f"{r['Итого_3мес_₴']:.2f}"
            w.writerow(r)

    # 6. pl_without_dds.csv — PL без шапочной ДДС или с пустой ТЧ или с несоответствием
    pl_problems = []
    for pl_uuid, pl in pl_items.items():
        проблемы = []
        if not pl["шапка_dds"]:
            проблемы.append("пустая_шапка_ДДС")
        if not pl["тч_dds"]:
            проблемы.append("пустая_ТЧ_Статьи")
        if pl["шапка_dds"] and pl["тч_dds"]:
            if pl["шапка_dds"] not in [t["uuid"] for t in pl["тч_dds"]]:
                проблемы.append("шапка_не_в_ТЧ")
        if проблемы:
            pl_problems.append({
                "PL_Код": pl["code"],
                "PL_Имя": pl["name"],
                "Группа": pl["group"],
                "ТипСтатьи": pl["тип"],
                "Проблемы": "; ".join(проблемы),
                "Авто": pl["авто"],
            })

    pl_path = config.REPORTS_DIR / "pl_without_dds.csv"
    with pl_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["PL_Код", "PL_Имя", "Группа", "ТипСтатьи",
                                         "Проблемы", "Авто"], delimiter=";")
        w.writeheader()
        for r in pl_problems:
            w.writerow(r)

    print(f"\n=== Результаты ===")
    print(f"Непокрытых ДДС: {len(uncovered_rows)}  → {uncov_path.name}")
    print(f"  сумма 3 месяца: {sum(u['Итого_3мес_₴'] if isinstance(u['Итого_3мес_₴'], (int, float)) else float(u['Итого_3мес_₴']) for u in uncovered_rows):,.2f} ₴")
    print(f"ДДС-дублей: {len(dup_rows)}  → {dup_path.name}")
    print(f"PL с проблемами: {len(pl_problems)}  → {pl_path.name}")


if __name__ == "__main__":
    main()
