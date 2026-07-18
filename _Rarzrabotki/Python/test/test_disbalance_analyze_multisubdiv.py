# -*- coding: utf-8 -*-
"""
Тест 1: анализ дисбаланса по подразделениям для двух целевых документов.

Воспроизводит запрос обработки А_ОбработкаДисбалансаПоПодразделениям
из ObjectModule.bsl:25-46 для:
  - Приобретение товаров и услуг ІБ00-008058 от 31.12.2025 11:00:00
  - Приобретение товаров и услуг ІБ00-000597 от 10.02.2026 11:00:00

Проверяет:
  - Сумма(Контроль) ≈ 0 по организации
  - Количество подразделений с Контроль <> 0 (ожидаем ≥ 2)
  - Если ≥ 3 — нужен жадный алгоритм (N пар)
"""
import sys
import win32com.client
import datetime

# Force UTF-8 output
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

v8 = win32com.client.Dispatch("V83.COMConnector")
CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

print("Connecting to BaseERP...")
conn = v8.Connect(CONN_ERP)
print("Connected OK\n")


def find_doc_by_uuid(uuid_str):
    """Находит документ ПриобретениеТоваровУслуг по UUID."""
    q = conn.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ ПЕРВЫЕ 1
        Т.Ссылка КАК Ссылка,
        Т.Номер КАК Номер,
        Т.Дата КАК Дата,
        Т.Проведен КАК Проведен,
        Т.Организация КАК Организация,
        Т.СуммаДокумента КАК СуммаДокумента
    ИЗ Документ.ПриобретениеТоваровУслуг КАК Т
    ГДЕ Т.Ссылка = &Ссылка
    """
    uid = conn.NewObject("УникальныйИдентификатор", uuid_str)
    ref = conn.Документы.ПриобретениеТоваровУслуг.ПолучитьСсылку(uid)
    q.УстановитьПараметр("Ссылка", ref)
    sel = q.Выполнить().Выбрать()
    if sel.Следующий():
        return sel.Ссылка, {
            "Номер": sel.Номер,
            "Дата": sel.Дата,
            "Проведен": sel.Проведен,
            "Организация": sel.Организация,
            "СуммаДокумента": sel.СуммаДокумента,
        }
    return None, None


def analyze_document(uuid_str, period_start, period_end, label):
    print("=" * 70)
    print(f"Анализ: {label}")
    print("=" * 70)

    doc_ref, info = find_doc_by_uuid(uuid_str)
    if doc_ref is None:
        print(f"  ДОКУМЕНТ НЕ НАЙДЕН по UUID {uuid_str}")
        return None
    print(f"  Найден: Номер={info['Номер']}, Дата={info['Дата']}, Проведен={info['Проведен']}, СуммаДок={info['СуммаДокумента']:.2f}, Орг={info['Организация']}")

    q_turn = conn.NewObject("Запрос")
    q_turn.Text = """
    ВЫБРАТЬ
        Об.Регистратор КАК Документ,
        Об.Подразделение КАК Подразделение,
        Об.Организация КАК Организация,
        Об.Статья КАК Статья,
        Об.СуммаПриход КАК Дебет,
        Об.СуммаРасход КАК Кредит,
        Об.СуммаПриход - Об.СуммаРасход КАК Контроль
    ИЗ
        РегистрНакопления.ПрочиеАктивыПассивы.Обороты(
            &НачПериода, &КонПериода, Регистратор, ) КАК Об
    ГДЕ
        Об.Регистратор = &Регистратор
        И НЕ Об.Статья В (
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВложенияСобственныхСредств),
            ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ВыводСобственныхСредств))
    """
    q_turn.УстановитьПараметр("Регистратор", doc_ref)
    q_turn.УстановитьПараметр("НачПериода", datetime.datetime.fromisoformat(period_start))
    q_turn.УстановитьПараметр("КонПериода", datetime.datetime.fromisoformat(period_end))

    table = q_turn.Выполнить().Выгрузить()
    print(f"\n  Обороты: {table.Количество()} строк")

    if table.Количество() == 0:
        print("  Нет движений в регистре ПрочиеАктивыПассивы по этому регистратору.")
        print("  Возможные причины: документ не провёл движения; расчёт себестоимости ещё не сделан.")
        return {"doc_ref": doc_ref, "podr_map": {}, "total": 0.0, "disb_count": 0}

    by_podr = {}
    statiyi = {}  # статьи с суммой по подразделению
    for i in range(table.Количество()):
        row = table.Получить(i)
        key = conn.XMLСтрока(row.Подразделение)
        name = conn.String(row.Подразделение)
        if key not in by_podr:
            by_podr[key] = {"Подразделение": row.Подразделение, "Имя": name, "Контроль": 0.0}
        by_podr[key]["Контроль"] += float(row.Контроль)

        sk = (key, conn.String(row.Статья))
        if sk not in statiyi:
            statiyi[sk] = {"ПодрИмя": name, "СтатьяИмя": conn.String(row.Статья), "Контроль": 0.0}
        statiyi[sk]["Контроль"] += float(row.Контроль)

    print(f"\n  Дисбаланс по подразделениям:")
    total = 0.0
    disb_count = 0
    for key, v in sorted(by_podr.items(), key=lambda kv: -kv[1]["Контроль"]):
        total += v["Контроль"]
        if round(v["Контроль"], 2) != 0:
            disb_count += 1
        marker = "+Дт" if round(v["Контроль"], 2) > 0 else ("-Кт" if round(v["Контроль"], 2) < 0 else " = ")
        print(f"    {marker} {v['Имя']:<50s} Контроль={v['Контроль']:>15.2f}")

    print(f"\n  По статьям внутри подразделений (первые 15 по |Контроль|):")
    sorted_st = sorted(statiyi.items(), key=lambda kv: -abs(kv[1]["Контроль"]))[:15]
    for (pk, st), v in sorted_st:
        if round(v["Контроль"], 2) != 0:
            print(f"    {v['ПодрИмя']:<40s} | {v['СтатьяИмя']:<40s} | {v['Контроль']:>15.2f}")

    print(f"\n  Итого Контроль по организации: {total:.2f} (должно быть ≈ 0)")
    print(f"  Подразделений с дисбалансом: {disb_count}")
    if round(total, 2) != 0:
        print(f"  [!] ПРЕДУПРЕЖДЕНИЕ: баланс по организации не сходится!")
    if disb_count >= 3:
        print(f"  [+] {disb_count} подразделений с дисбалансом — текущая обработка НЕ справится (нужен N пар)")
    elif disb_count == 2:
        print(f"  [=] 2 подразделения — одна пара")
    else:
        print(f"  [-] < 2 подразделений — взаимозачёт не нужен")

    return {"doc_ref": doc_ref, "podr_map": by_podr, "total": total, "disb_count": disb_count}


# UUIDs из MCP find_document_ref
result1 = analyze_document(
    uuid_str="20b3ae2a-f767-11f0-8105-00155dce3d04",
    period_start="2025-12-01T00:00:00",
    period_end="2025-12-31T23:59:59",
    label="ІБ00-008058 от 31.12.2025 11:00:00",
)

print()
result2 = analyze_document(
    uuid_str="3732c824-08ad-11f1-8106-00155dce3d04",
    period_start="2026-02-01T00:00:00",
    period_end="2026-02-28T23:59:59",
    label="ІБ00-000597 от 10.02.2026 11:00:00 (на 66 245,81)",
)

print("\n" + "=" * 70)
print("ИТОГ:")
print("=" * 70)
for name, r in [("ІБ00-008058", result1), ("ІБ00-000597", result2)]:
    if r is None:
        print(f"  {name}: ДОКУМЕНТ НЕ НАЙДЕН")
    else:
        print(f"  {name}: подразделений с дисбалансом = {r['disb_count']}, Контроль = {r['total']:.2f}")
