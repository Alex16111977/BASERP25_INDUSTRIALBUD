# -*- coding: utf-8 -*-
"""
Заповнення документа КорректировкаРегистров 0Ц-00000001 від 30.11.2025
залишками з Казни по регістру А_ДенежныеСредстваФинАгенты.

Алгоритм:
1. Беремо залишки (Остатки) з Казни на 01.12.2025
2. Маппінг: ФинАгент(по Коду), Направление(по Коду), ДокументПередачи(по UUID)
3. Додаткові виміри: Подразделение, Контрагент з документа ДокПередачи;
   ДоговорКонтрагента, ОбъектРасчетов з РасшифровкаПлатежа документа ДокПередачи
4. Ресурси: Сумма=Казна, СуммаСКомиссией=Сумма/(1-Процент/100),
   СуммаКомиссии=СуммаСКомиссией*Процент/100
5. Записуємо рухи (Приход) в документ корекції

Режим DRY_RUN = True — тільки вивід без запису.
Змінити на False для реального запису.
"""
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime

# ============================================================
# НАЛАШТУВАННЯ
# ============================================================
DRY_RUN = True  # True = тільки вивід, False = запис в базу

CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
CONN_KAZNA = 'Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"'

CORRECTION_DOC_NUMBER = "0Ц-00000001"
CORRECTION_DATE = datetime(2025, 11, 30, 23, 59, 59)
BALANCE_DATE = datetime(2025, 12, 1)  # Остатки на початок 01.12 = кінець 30.11


def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")

    print("Підключення до Казни...")
    kazna = v8.Connect(CONN_KAZNA)
    print("OK")

    print("Підключення до ЕРП...")
    erp = v8.Connect(CONN_ERP)
    print("OK\n")

    # ============================================================
    # КРОК 1: Залишки з Казни
    # ============================================================
    print("=" * 80)
    print("КРОК 1: Отримання залишків з Казни на", BALANCE_DATE.strftime("%d.%m.%Y"))
    print("=" * 80)

    q = kazna.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ "
        "  ДС.ФинАгент КАК ФинАгент, "
        "  ДС.Направление КАК Направление, "
        "  ДС.ДокументПередачи КАК ДокументПередачи, "
        "  ДС.СуммаОстаток КАК СуммаОстаток "
        "ИЗ "
        "  РегистрНакопления.ДенежныеСредстваФинАгенты.Остатки(&Дата) КАК ДС "
        "ГДЕ ДС.СуммаОстаток <> 0"
    )
    q.УстановитьПараметр("Дата", pywintypes.Time(BALANCE_DATE))

    res = q.Выполнить()
    sel = res.Выбрать()

    kazna_rows = []
    total_kazna = 0

    while sel.Следующий():
        row = {}
        row["summa_kazna"] = float(sel.СуммаОстаток)
        total_kazna += row["summa_kazna"]

        # ФинАгент
        row["fa_kazna_name"] = kazna.String(sel.ФинАгент) if kazna.ЗначениеЗаполнено(sel.ФинАгент) else ""
        row["fa_kazna_code"] = ""
        try:
            if kazna.ЗначениеЗаполнено(sel.ФинАгент):
                row["fa_kazna_code"] = str(sel.ФинАгент.Код).strip()
        except:
            pass

        # Направление
        row["napr_kazna_name"] = ""
        row["napr_kazna_code"] = ""
        try:
            if kazna.ЗначениеЗаполнено(sel.Направление):
                row["napr_kazna_name"] = kazna.String(sel.Направление)
                row["napr_kazna_code"] = str(sel.Направление.Код).strip()
        except:
            pass

        # ДокументПередачи — UUID
        row["doc_peredachi_uuid"] = ""
        row["doc_peredachi_type"] = ""
        row["doc_peredachi_filled"] = False
        try:
            if kazna.ЗначениеЗаполнено(sel.ДокументПередачи):
                row["doc_peredachi_filled"] = True
                row["doc_peredachi_uuid"] = kazna.XMLСтрока(
                    sel.ДокументПередачи.УникальныйИдентификатор()
                )
                row["doc_peredachi_type"] = kazna.String(
                    sel.ДокументПередачи.Метаданные().Имя
                )
        except Exception as e:
            print(f"  WARN: Помилка UUID ДокПередачи: {e}")

        kazna_rows.append(row)

    print(f"  Всього рядків з Казни: {len(kazna_rows)}")
    print(f"  Загальна сума: {total_kazna:,.2f}\n")

    # ============================================================
    # КРОК 2: Маппінг Казна -> ЕРП
    # ============================================================
    print("=" * 80)
    print("КРОК 2: Маппінг Казна -> ЕРП")
    print("=" * 80)

    movements_to_write = []
    mapping_errors = []

    for i, row in enumerate(kazna_rows):
        entry = {}
        entry["summa"] = row["summa_kazna"]
        entry["fa_kazna"] = row["fa_kazna_name"]
        entry["napr_kazna"] = row["napr_kazna_name"]

        # 2a. ФинАгент — по Коду
        entry["fin_agent_erp"] = None
        if row["fa_kazna_code"]:
            try:
                fa_ref = erp.Справочники.А_ФинАгенты.НайтиПоКоду(row["fa_kazna_code"])
                if erp.ЗначениеЗаполнено(fa_ref):
                    entry["fin_agent_erp"] = fa_ref
                else:
                    mapping_errors.append(f"Рядок {i+1}: ФинАгент '{row['fa_kazna_name']}' код={row['fa_kazna_code']} — НЕ знайдено в ЕРП")
            except Exception as e:
                mapping_errors.append(f"Рядок {i+1}: Помилка маппінгу ФинАгент: {e}")

        if not entry["fin_agent_erp"]:
            mapping_errors.append(f"Рядок {i+1}: ПРОПУСК — ФинАгент '{row['fa_kazna_name']}' не знайдено")
            continue

        # 2b. НаправлениеДеятельности — по Коду
        entry["direction_erp"] = None
        if row["napr_kazna_code"]:
            try:
                napr_ref = erp.Справочники.НаправленияДеятельности.НайтиПоКоду(row["napr_kazna_code"])
                if erp.ЗначениеЗаполнено(napr_ref):
                    entry["direction_erp"] = napr_ref
            except:
                pass
        # Якщо не знайшли — пуста ссилка (допустимо)
        if not entry["direction_erp"]:
            entry["direction_erp"] = erp.Справочники.НаправленияДеятельности.ПустаяСсылка()

        # 2c. ДокументПередачи — по UUID
        entry["doc_peredachi_erp"] = erp.Документы.СписаниеБезналичныхДенежныхСредств.ПустаяСсылка()
        entry["percent_fa"] = 0.0
        entry["subdivision"] = erp.Справочники.СтруктураПредприятия.ПустаяСсылка()
        entry["counterparty"] = erp.Справочники.Контрагенты.ПустаяСсылка()
        entry["contract"] = erp.Справочники.ДоговорыКонтрагентов.ПустаяСсылка()
        entry["settlement_obj"] = erp.Справочники.ОбъектыРасчетов.ПустаяСсылка()

        if row["doc_peredachi_filled"] and row["doc_peredachi_uuid"]:
            try:
                uid = erp.NewObject("УникальныйИдентификатор", row["doc_peredachi_uuid"])
                doc_ref = erp.Документы.СписаниеБезналичныхДенежныхСредств.ПолучитьСсылку(uid)
                entry["doc_peredachi_erp"] = doc_ref

                # Витягнути додаткові поля з документа
                try:
                    entry["subdivision"] = doc_ref.Подразделение
                    entry["counterparty"] = doc_ref.Контрагент
                    # А_ПроцентФинАгента
                    pct = float(doc_ref.А_ПроцентФинАгента)
                    entry["percent_fa"] = pct
                except Exception as e:
                    # Документ може не існувати (ПолучитьСсылку створює "фантомну" ссилку)
                    pass

                # Витягнути ДоговорКонтрагента та ОбъектРасчетов
                # з РасшифровкаПлатежа документа ДокументПередачи
                # (згідно коду проведення ПровестиРегистрА_ДенежныеСредстваФинАгенты)
                try:
                    rp = doc_ref.РасшифровкаПлатежа
                    rp_count = rp.Количество()
                    for rp_idx in range(rp_count):
                        rp_row = rp.Получить(rp_idx)
                        or_val = rp_row.ОбъектРасчетов
                        if erp.ЗначениеЗаполнено(or_val):
                            entry["settlement_obj"] = or_val
                            # ДоговорКонтрагента = ОбъектРасчетов.Договор
                            try:
                                dog = or_val.Договор
                                if erp.ЗначениеЗаполнено(dog):
                                    entry["contract"] = dog
                            except:
                                pass
                            break  # беремо перший рядок з заповненим ОбъектРасчетов
                except:
                    pass
            except Exception as e:
                print(f"  WARN: Помилка маппінгу ДокПередачи UUID={row['doc_peredachi_uuid']}: {e}")

        # Якщо Подразделение пусте — зворотній маппінг через НаправлениеДеятельности
        if not erp.ЗначениеЗаполнено(entry["subdivision"]) and erp.ЗначениеЗаполнено(entry["direction_erp"]):
            try:
                q_sub = erp.NewObject("Запрос")
                q_sub.Текст = (
                    "ВЫБРАТЬ ПЕРВЫЕ 1 "
                    "  П.Ссылка КАК Подразделение "
                    "ИЗ Справочник.СтруктураПредприятия КАК П "
                    "ГДЕ П.А_НаправлениеДеятельности = &Направление "
                    "  И НЕ П.ПометкаУдаления"
                )
                q_sub.УстановитьПараметр("Направление", entry["direction_erp"])
                res_sub = q_sub.Выполнить()
                if not res_sub.Пустой():
                    sel_sub = res_sub.Выбрать()
                    sel_sub.Следующий()
                    entry["subdivision"] = sel_sub.Подразделение
            except:
                pass

        # КРОК 4: Розрахувати ресурси
        summa = entry["summa"]
        pct = entry["percent_fa"]
        if pct > 0 and pct < 100:
            summa_s_komissiei = round(summa / (1.0 - pct / 100.0), 2)
            summa_komissii = round(summa_s_komissiei * pct / 100.0, 2)
        else:
            summa_s_komissiei = summa
            summa_komissii = 0.0

        entry["summa_komissii"] = summa_komissii
        entry["summa_s_komissiei"] = summa_s_komissiei

        movements_to_write.append(entry)

    # ============================================================
    # ВИВІД РЕЗУЛЬТАТІВ
    # ============================================================
    print(f"\n  Успішно замаплено: {len(movements_to_write)} рядків")
    if mapping_errors:
        print(f"  Помилки маппінгу: {len(mapping_errors)}")
        for err in mapping_errors:
            print(f"    WARNING: {err}")

    print("\n" + "=" * 80)
    print("КРОК 3: Деталізація рухів для запису")
    print("=" * 80)

    total_summa = 0
    total_komissii = 0
    total_s_komissiei = 0

    for i, e in enumerate(movements_to_write):
        fa_name = erp.String(e["fin_agent_erp"]) if e["fin_agent_erp"] else "?"
        napr_name = erp.String(e["direction_erp"]) if erp.ЗначениеЗаполнено(e["direction_erp"]) else "<пусто>"
        doc_name = erp.String(e["doc_peredachi_erp"]) if erp.ЗначениеЗаполнено(e["doc_peredachi_erp"]) else "<пусто>"
        sub_name = erp.String(e["subdivision"]) if erp.ЗначениеЗаполнено(e["subdivision"]) else "<пусто>"
        ctr_name = erp.String(e["counterparty"]) if erp.ЗначениеЗаполнено(e["counterparty"]) else "<пусто>"

        total_summa += e["summa"]
        total_komissii += e["summa_komissii"]
        total_s_komissiei += e["summa_s_komissiei"]

        print(f"\n  [{i+1}] ФинАгент: {fa_name}")
        print(f"      Направление: {napr_name}")
        print(f"      ДокПередачи: {doc_name}")
        print(f"      Подразделение: {sub_name}")
        print(f"      Контрагент: {ctr_name}")
        print(f"      %ФА: {e['percent_fa']}")
        print(f"      Сумма(Казна):      {e['summa']:>15,.2f}")
        print(f"      СуммаКомиссии:     {e['summa_komissii']:>15,.2f}")
        print(f"      СуммаСКомиссией:   {e['summa_s_komissiei']:>15,.2f}")

    print(f"\n  {'-' * 50}")
    print(f"  РАЗОМ Сумма:            {total_summa:>15,.2f}")
    print(f"  РАЗОМ СуммаКомиссии:    {total_komissii:>15,.2f}")
    print(f"  РАЗОМ СуммаСКомиссией:  {total_s_komissiei:>15,.2f}")

    # ============================================================
    # КРОК 5: Запис в документ
    # ============================================================
    if DRY_RUN:
        print(f"\n{'=' * 80}")
        print("DRY RUN — запис НЕ виконано. Змініть DRY_RUN = False для запису.")
        print(f"{'=' * 80}")
        return

    print(f"\n{'=' * 80}")
    print("КРОК 5: Запис рухів в КорректировкаРегистров")
    print(f"{'=' * 80}")

    # Знайти документ
    doc_ref = erp.Документы.КорректировкаРегистров.НайтиПоНомеру(
        CORRECTION_DOC_NUMBER,
        pywintypes.Time(CORRECTION_DATE)
    )

    if not erp.ЗначениеЗаполнено(doc_ref):
        print(f"  ПОМИЛКА: Документ {CORRECTION_DOC_NUMBER} не знайдено!")
        return

    print(f"  Знайдено документ: {erp.String(doc_ref)}")

    doc_obj = doc_ref.ПолучитьОбъект()

    # Прочитати поточні рухи, очистити, додати нові
    movements = doc_obj.Движения.А_ДенежныеСредстваФинАгенты
    movements.Прочитать()
    print(f"  Поточних рухів: {movements.Количество()}")
    movements.Очистить()

    period = pywintypes.Time(CORRECTION_DATE)

    for e in movements_to_write:
        m = movements.Добавить()
        m.ВидДвижения = erp.ВидДвиженияНакопления.Приход
        m.Период = period
        m.Активность = True
        m.ФинАгент = e["fin_agent_erp"]
        m.ДокументПередачи = e["doc_peredachi_erp"]
        m.НаправлениеДеятельности = e["direction_erp"]
        m.Подразделение = e["subdivision"]
        m.Контрагент = e["counterparty"]
        m.ДоговорКонтрагента = e["contract"]
        m.ОбъектРасчетов = e["settlement_obj"]
        m.Сумма = e["summa"]
        m.СуммаКомиссии = e["summa_komissii"]
        m.СуммаСКомиссией = e["summa_s_komissiei"]

    print(f"  Додано рухів: {movements.Количество()}")

    # Записати рухи напряму
    try:
        movements.Записать()
        print(f"  OK: movements.Записать() — OK")
    except Exception as e:
        print(f"  FAIL: movements.Записать() ПОМИЛКА: {e}")
        # Альтернативний підхід — записати через документ
        print("  Спроба через doc_obj.Записать()...")
        try:
            doc_obj.Движения.А_ДенежныеСредстваФинАгенты.Записывать = True
            doc_obj.ОбменДанными.Загрузка = True
            doc_obj.Записать()
            print(f"  OK: doc_obj.Записать() — OK")
        except Exception as e2:
            print(f"  FAIL: doc_obj.Записать() ПОМИЛКА: {e2}")
            return

    # Перевірка — прочитати назад
    print("\n  Перевірка запису...")
    movements2 = doc_obj.Движения.А_ДенежныеСредстваФинАгенты
    movements2.Прочитать()
    print(f"  Рухів в документі після запису: {movements2.Количество()}")

    print("\n=== ГОТОВО ===")


if __name__ == "__main__":
    main()
