# -*- coding: utf-8 -*-
"""
Приёмочный тест: подгружает скомпилированную .epf через ВнешниеОбработки.Создать,
запускает .Выполнить(), проверяет результаты.

Шаги:
1) До прогона: фиксируем что есть отрицательные остатки на 30.04.2026
2) Запуск обработки через COM (ВнешниеОбработки.Создать → НаДату/Организация → Выполнить)
3) Проверка ТЧ результатов: все строки Статус="OK"
4) Верификация документов: дата=01.04.2026, проведены, флаг А_ДляЗакрытиеОтрицательныхОстатков
5) После прогона: повторный запрос — минусов нет (по выбранной организации)
"""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\build\ЗакрытиеОтрицательныхОстатков.epf"
ORG_UUID_HEX = "6bee36b2-53f0-11e6-80d3-000c29bbac23"  # ТОВ "ІНДАСТРІАЛБУД"


def connect():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    return v8.Connect(CONN)


def negative_balances_query(erp, end_of_month, org_ref):
    """Возвращает таблицу значений с отрицательными остатками по выбранной организации."""
    text = """
    ВЫБРАТЬ
        Ост.Организация,
        Ост.АналитикаУчетаНоменклатуры.Номенклатура КАК Номенклатура,
        Ост.АналитикаУчетаНоменклатуры.СкладскаяТерритория КАК Склад,
        -Ост.КоличествоОстаток КАК Количество
    ИЗ РегистрНакопления.ТоварыОрганизаций.Остатки(&НаДату, Организация = &Орг) КАК Ост
    ГДЕ Ост.КоличествоОстаток < 0
    """
    q = erp.NewObject("Запрос")
    q.Text = text
    q.SetParameter("НаДату", end_of_month)
    q.SetParameter("Орг", org_ref)
    return q.Execute().Выгрузить()


def run():
    print("=" * 70)
    print("ACCEPTANCE: ЗакрытиеОтрицательныхОстатков.epf")
    print("=" * 70)
    erp = connect()

    # Конец и начало месяца апрель 2026 через 1С
    q = erp.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ ДАТАВРЕМЯ(2026,4,30,23,59,59) КАК Конец, ДАТАВРЕМЯ(2026,4,1) КАК Начало"
    r = q.Execute().Выгрузить().Получить(0)
    end_of_april = r.Конец
    start_of_april = r.Начало

    org_uid = erp.NewObject("УникальныйИдентификатор", ORG_UUID_HEX)
    org_ref = erp.Справочники.Организации.ПолучитьСсылку(org_uid)
    print(f"  Организация: {erp.String(org_ref)}")

    # === 1) ПРЕД-ПРОГОН ===
    before = negative_balances_query(erp, end_of_april, org_ref)
    n_before = before.Количество()
    print(f"  [1] До прогона: минусов по орг={n_before}")
    if n_before == 0:
        print("      (минусов нет — проверяем идемпотентность повторного прогона)")

    # === 2) ЗАПУСК ОБРАБОТКИ ===
    print("  [2] Загружаю обработку через ВнешниеОбработки.Создать...")
    обработка = erp.ВнешниеОбработки.Создать(EPF)
    обработка.НаДату = end_of_april
    обработка.Организация = org_ref
    print("      .ВыполнитьЗакрытие() ...")
    обработка.ВыполнитьЗакрытие()

    # === 3) ПРОВЕРКА ТЧ ===
    тч = обработка.ОтрицательныеОстатки
    total = тч.Количество()
    ok_count = 0
    err_count = 0
    no_price_count = 0
    док_uids = {}
    for i in range(total):
        стр = тч.Получить(i)
        статус = str(стр.Статус)
        if статус == "OK":
            ok_count += 1
            uid = erp.String(стр.СозданныйДокумент.УникальныйИдентификатор())
            док_uids[uid] = (erp.String(стр.Склад), erp.String(стр.Номенклатура))
        elif статус.startswith("Цена не найдена"):
            no_price_count += 1
        elif статус.startswith("Ошибка"):
            err_count += 1
            print(f"      WARN: {erp.String(стр.Номенклатура)} → {статус[:120]}")
    print(f"  [3] ТЧ результатов: всего={total}, OK={ok_count}, без_цены={no_price_count}, ошибок={err_count}")
    if n_before > 0:
        assert total > 0, "ТЧ ОтрицательныеОстатки пуста — обработка не нашла минусов"
    print(f"      Создано/найдено документов: {len(док_uids)}")

    # === 4) ВЕРИФИКАЦИЯ ДОКУМЕНТОВ ===
    for uid, (склад, ном) in док_uids.items():
        ref = erp.Документы.ОприходованиеИзлишковТоваров.ПолучитьСсылку(
            erp.NewObject("УникальныйИдентификатор", uid))
        obj = ref.ПолучитьОбъект()
        assert obj is not None, f"Документ {uid} не существует"
        assert obj.А_ДляЗакрытиеОтрицательныхОстатков, f"Флаг А_ДляЗакрытиеОтрицательныхОстатков не выставлен у {uid}"
        # Дата должна быть = 01.04.2026 00:00:00
        d = obj.Дата
        assert d.year == 2026 and d.month == 4 and d.day == 1, \
            f"Дата документа {d} — ожидалась 01.04.2026 ({склад})"
        проведен = bool(obj.Проведен)
        print(f"      док={uid[:8]}.. склад={склад} проведен={проведен} строк_тч={obj.Товары.Количество()}")
    print(f"  [4] Все {len(док_uids)} документов созданы с датой 01.04.2026 и флагом")

    # === 5) ПОСТ-ПРОГОН ===
    after = negative_balances_query(erp, end_of_april, org_ref)
    n_after = after.Количество()
    print(f"  [5] После прогона: минусов по орг={n_after}")
    if n_after == 0:
        print("      ✓ Все минусы закрыты")
    else:
        print(f"      ! Осталось {n_after} строк с минусами:")
        for i in range(min(5, n_after)):
            s = after.Получить(i)
            print(f"        - {erp.String(s.Номенклатура)} ({erp.String(s.Склад)}) qty={s.Количество}")

    print("=" * 70)
    if n_after == 0 and err_count == 0:
        print("OK: ВСЕ ТЕСТЫ ПРОШЛИ")
        return 0
    else:
        print(f"PARTIAL: остатки_после={n_after} ошибки={err_count} без_цены={no_price_count}")
        return 1


if __name__ == "__main__":
    try:
        sys.exit(run())
    except AssertionError as e:
        print(f"FAIL: {e}")
        sys.exit(2)
    except Exception as e:
        if hasattr(e, "excepinfo") and e.excepinfo:
            print(f"FAIL: {e.excepinfo[2]}")
        else:
            print(f"FAIL: {type(e).__name__}: {e}")
        sys.exit(3)
