"""
Rashifrovka: проверить что ТЧ А_РасшифровкаВыплатыЗарплатаПоФизлицам
заполнена корректно по ТЧ Распределение Ф2 (после Свернуть).
"""
import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

CONN_ERP = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN_ERP)

    # Получаем эталон Ф2
    q = erp.NewObject("Запрос")
    q.Текст = '''ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка КАК Ссылка
ИЗ Документ.РаспределениеФ2 КАК Док
ГДЕ Док.Номер = "000000026" И Док.Проведен'''
    rs = q.Выполнить()
    sel = rs.Выбрать(); sel.Следующий()
    f2_ref = sel.Ссылка
    f2 = f2_ref.ПолучитьОбъект()

    # Перепровести (если не было)
    f2.Записать(erp.РежимЗаписиДокумента.Проведение)
    f2 = f2_ref.ПолучитьОбъект()

    vk = f2.А_ВедомостьВКассу.ПолучитьОбъект()

    # Σ ОтработаноЧасов в А_Расшифровке должна быть = Σ Распределение.ОтработаноЧасов
    sum_h_f2 = sum(float(f2.Распределение.Получить(i).ОтработаноЧасов)
                   for i in range(f2.Распределение.Количество()))
    sum_h_vk = sum(float(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Получить(i).ОтработаноЧасов)
                   for i in range(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Количество()))
    print(f"Σ ОтработаноЧасов: Ф2={sum_h_f2}, ВКассу={sum_h_vk}")
    if abs(sum_h_f2 - sum_h_vk) > 0.01:
        print(f"FAIL: Несовпадение часов: {sum_h_f2} != {sum_h_vk}")
        sys.exit(1)

    # Σ Суммы в А_Расшифровке должна быть = Σ Распределение.СуммаНачисления
    sum_s_f2 = sum(float(f2.Распределение.Получить(i).СуммаНачисления)
                   for i in range(f2.Распределение.Количество()))
    sum_s_vk = sum(float(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Получить(i).Сумма)
                   for i in range(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Количество()))
    print(f"Σ Сумма: Ф2={sum_s_f2:,.2f}, ВКассу={sum_s_vk:,.2f}")
    if abs(sum_s_f2 - sum_s_vk) > 0.01:
        print(f"FAIL: Несовпадение сумм: {sum_s_f2} != {sum_s_vk}")
        sys.exit(1)

    # У каждой строки Расшифровки должны быть заполнены ФизЛицо, Подр
    cnt = vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Количество()
    warn_no_statya = 0
    fail_no_fl_or_podr = 0
    for i in range(cnt):
        r = vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Получить(i)
        if r.ФизическоеЛицо.Пустая():
            print(f"FAIL: строка {i+1}: ФизическоеЛицо пусто")
            fail_no_fl_or_podr += 1
        if r.Подразделение.Пустая():
            print(f"FAIL: строка {i+1}: Подразделение пусто")
            fail_no_fl_or_podr += 1
        if r.СтатьяДвиженияДенежныхСредств.Пустая():
            warn_no_statya += 1

    if fail_no_fl_or_podr > 0:
        sys.exit(1)

    if warn_no_statya > 0:
        print(f"WARN: {warn_no_statya} строк без СтатьиДДС (у сотрудника не заполнен реквизит СтатьяЗарплата)")

    print(f"\nPASS: rashifrovka test пройден ({cnt} строк)")


if __name__ == "__main__":
    main()
