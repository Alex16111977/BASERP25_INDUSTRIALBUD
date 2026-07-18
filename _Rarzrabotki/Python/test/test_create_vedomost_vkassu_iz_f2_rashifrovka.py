"""
Rashifrovka: проверить что ТЧ А_РасшифровкаВыплатыЗарплатаПоФизлицам
заполнена корректно — зеркало ТЧ.Зарплата 1:1 (после унификации 2026-05-27).
"""
import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

import win32com.client

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


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

    # Кол строк А_Расшифровки = кол строк ТЧ.Зарплата (зеркало 1:1)
    cnt_rash = vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Количество()
    cnt_zarp = vk.Зарплата.Количество()
    print(f"Кол строк: ТЧ.Зарплата={cnt_zarp}, А_Расшифровка={cnt_rash}")
    if cnt_rash != cnt_zarp:
        print(f"FAIL: А_Расшифровка должна зеркалировать Зарплату 1:1 ({cnt_zarp} != {cnt_rash})")
        sys.exit(1)

    # Σ КВыплате в А_Расшифровке = Σ КВыплате в ТЧ.Зарплата = Σ Распределение Ф2.СуммаНачисления
    sum_s_f2 = sum(float(f2.Распределение.Получить(i).СуммаНачисления)
                   for i in range(f2.Распределение.Количество()))
    sum_s_zarp = sum(float(vk.Зарплата.Получить(i).КВыплате)
                     for i in range(vk.Зарплата.Количество()))
    sum_s_rash = sum(float(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Получить(i).КВыплате)
                     for i in range(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Количество()))
    print(f"Σ КВыплате: Ф2.Распределение={sum_s_f2:,.2f}, ТЧ.Зарплата={sum_s_zarp:,.2f}, А_Расшифровка={sum_s_rash:,.2f}")
    if abs(sum_s_f2 - sum_s_rash) > 0.01:
        print(f"FAIL: Несовпадение сумм Ф2 vs А_Расшифровка: {sum_s_f2} != {sum_s_rash}")
        sys.exit(1)
    if abs(sum_s_zarp - sum_s_rash) > 0.01:
        print(f"FAIL: Несовпадение сумм Зарплата vs А_Расшифровка: {sum_s_zarp} != {sum_s_rash}")
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
