# -*- coding: utf-8 -*-
"""
Діагностика: чому forward check (ОбробитиДокументиЕРП) показує 207 різниць
і 0 переказів для ТАС_Будівн за грудень 2025.
"""

import win32com.client
import pythoncom
import pywintypes
from datetime import datetime


def main():
    pythoncom.CoInitialize()
    v8 = win32com.client.Dispatch("V83.COMConnector")

    conn_erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
    conn_buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

    dt_start = pywintypes.Time(datetime(2025, 12, 1))
    dt_end = pywintypes.Time(datetime(2025, 12, 31, 23, 59, 59))

    # Get bank account
    q = conn_erp.NewObject("Запрос")
    q.Текст = (
        "ВЫБРАТЬ Ссылка, Наименование, НомерСчета "
        "ИЗ Справочник.БанковскиеСчетаОрганизаций "
        "ГДЕ Наименование ПОДОБНО &Имя И Наименование НЕ ПОДОБНО \"%$$$%\""
    )
    q.УстановитьПараметр("Имя", "%ТАС_Будівн_%Індастріалбуд%")
    sel = q.Выполнить().Выбрать()
    sel.Следующий()
    bank_ref = sel.Ссылка

    # ERP query
    q2 = conn_erp.NewObject("Запрос")
    q2.Текст = (
        "ВЫБРАТЬ "
        "  Об.Регистратор КАК Регистратор, "
        "  ПРЕДСТАВЛЕНИЕ(Об.Регистратор) КАК РегистраторПредст, "
        "  Об.СуммаПриход КАК Приход, "
        "  Об.СуммаРасход КАК Расход, "
        "  Соотв.УникальныйИдентификаторПриемника КАК UIDвBuhBud, "
        "  Соотв.ТипПриемника КАК ТипПриемника "
        "ИЗ "
        "  РегистрНакопления.ДенежныеСредстваБезналичные.Обороты("
        "    &НачалоПериода, &КонецПериода, Авто, "
        "    БанковскийСчет = &БанковскийСчет) КАК Об "
        "  ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.СоответствияОбъектовИнформационныхБаз КАК Соотв "
        "  ПО Об.Регистратор = Соотв.УникальныйИдентификаторИсточника "
        "    И (Соотв.УзелИнформационнойБазы ССЫЛКА ПланОбмена.ОбменУправлениеПредприятиемБухгалтерия20)"
    )
    q2.УстановитьПараметр("НачалоПериода", dt_start)
    q2.УстановитьПараметр("КонецПериода", dt_end)
    q2.УстановитьПараметр("БанковскийСчет", bank_ref)

    sel2 = q2.Выполнить().Выбрать()

    print("=" * 90)
    print("ДІАГНОСТИКА forward check: перші 5 документів")
    print("=" * 90)

    count = 0
    sample_ok = 0
    sample_diff = 0
    sample_transfer = 0
    sample_no_buh = 0

    while sel2.Следующий():
        count += 1
        uid_erp = str(conn_erp.XMLСтрока(sel2.Регистратор.УникальныйИдентификатор())).upper()
        prihod = float(sel2.Приход) if sel2.Приход else 0
        rashod = float(sel2.Расход) if sel2.Расход else 0
        net_erp = prihod - rashod
        predst = str(sel2.РегистраторПредст)

        # Check transfer
        is_transfer = False
        hoz_op_str = ""
        try:
            ho = sel2.Регистратор.ХозяйственнаяОперация
            hoz_op_str = str(conn_erp.XMLСтрока(ho))
            if "ДругогоСчета" in hoz_op_str or "ДругойСчет" in hoz_op_str:
                is_transfer = True
        except Exception as e:
            hoz_op_str = f"ПОМИЛКА: {e}"

        # Check BuhBud match
        uid_buh = ""
        typ_priem = ""
        try:
            if conn_erp.ЗначениеЗаполнено(sel2.UIDвBuhBud):
                uid_buh = str(sel2.UIDвBuhBud).strip()
            if conn_erp.ЗначениеЗаполнено(sel2.ТипПриемника):
                typ_priem = str(sel2.ТипПриемника).strip()
        except:
            pass

        net_buh = "?"
        status = "?"

        if is_transfer:
            status = "ПЕРЕКАЗ"
            sample_transfer += 1
        elif uid_buh and typ_priem:
            # Query BuhBud
            try:
                typ_clean = typ_priem.replace("ДокументСсылка.", "")
                buh_doc_ref = conn_buh.Документы[typ_clean].ПолучитьСсылку(
                    conn_buh.NewObject("УникальныйИдентификатор", uid_buh))
                q_buh = conn_buh.NewObject("Запрос")
                q_buh.Текст = (
                    "ВЫБРАТЬ СУММА(Т.Приход) - СУММА(Т.Расход) КАК Сумма "
                    "ИЗ ("
                    "  ВЫБРАТЬ Д.Сумма КАК Приход, 0 КАК Расход "
                    "  ИЗ РегистрБухгалтерии.Хозрасчетный КАК Д "
                    "  ГДЕ Д.Регистратор = &Регистратор "
                    "    И Д.СчетДт В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках)) "
                    "  ОБЪЕДИНИТЬ ВСЕ "
                    "  ВЫБРАТЬ 0, Д.Сумма "
                    "  ИЗ РегистрБухгалтерии.Хозрасчетный КАК Д "
                    "  ГДЕ Д.Регистратор = &Регистратор "
                    "    И Д.СчетКт В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках))) КАК Т"
                )
                q_buh.УстановитьПараметр("Регистратор", buh_doc_ref)
                sel_buh = q_buh.Выполнить().Выбрать()
                if sel_buh.Следующий():
                    net_buh = float(sel_buh.Сумма) if sel_buh.Сумма else 0
                else:
                    net_buh = 0

                if abs(round(net_erp - net_buh, 2)) == 0:
                    status = "ОК"
                    sample_ok += 1
                else:
                    status = f"РІЗНИЦЯ (ERP={net_erp:,.2f}, BuhBud={net_buh:,.2f})"
                    sample_diff += 1
            except Exception as e:
                status = f"ПОМИЛКА BuhBud: {e}"
                sample_diff += 1
        else:
            status = "НЕТ ВІДПОВІДНОСТІ"
            sample_no_buh += 1

        if count <= 10 or (isinstance(status, str) and "РІЗНИЦЯ" in status and sample_diff <= 3):
            print(f"\n  #{count}: {predst[:60]}")
            print(f"    ХозОперація: {hoz_op_str[:60]}")
            print(f"    UIDвBuhBud: '{uid_buh}', Тип: '{typ_priem}'")
            print(f"    ERP нетто: {net_erp:,.2f}")
            print(f"    BuhBud нетто: {net_buh}")
            print(f"    Статус: {status}")

    print(f"\n{'='*90}")
    print(f"  Всього: {count}")
    print(f"    ОК: {sample_ok}")
    print(f"    Переказ: {sample_transfer}")
    print(f"    Без відповідності: {sample_no_buh}")
    print(f"    Різниця: {sample_diff}")

    pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
