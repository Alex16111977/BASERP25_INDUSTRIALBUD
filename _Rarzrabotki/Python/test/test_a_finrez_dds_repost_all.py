"""Repost усіх документів А_ФинРез_DDS після рефактору ОбработкаПроведения під нову структуру А_ОтчетDDS_Свод.

Знаходить непомічені на видалення документи і виконує перепроведення.
Якщо проведення FAIL — друкує помилку для fix-loop.

Очікувано: ok = N, fail = 0 (де N = число всіх документів).
"""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def main():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN)

    q = erp.NewObject("Запрос")
    q.Text = """
        ВЫБРАТЬ
            Ссылка КАК Ref,
            Номер  КАК Num,
            Дата   КАК Dt,
            Месяц  КАК Mn,
            Проведен КАК Posted
        ИЗ Документ.А_ФинРез_DDS
        ГДЕ НЕ ПометкаУдаления
        УПОРЯДОЧИТЬ ПО Месяц
    """
    rows = q.Execute().Выгрузить()
    total = rows.Количество()
    print(f"Знайдено документів А_ФинРез_DDS (без помітки на видалення): {total}")

    ok, fail = 0, 0
    failures = []
    for i in range(total):
        row = rows.Получить(i)
        ref, num, dt, mn, was = row.Ref, row.Num, row.Dt, row.Mn, row.Posted
        try:
            obj = ref.ПолучитьОбъект()
            if was:
                obj.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
                obj = ref.ПолучитьОбъект()
            obj.Записать(erp.РежимЗаписиДокумента.Проведение)
            print(f">>> {num} | {mn:%m.%Y} | was_posted={was} - Posted OK")
            ok += 1
        except Exception as e:
            err = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
            print(f">>> {num} | {mn:%m.%Y} - FAIL: {err}")
            failures.append((num, mn, err))
            fail += 1

    print(f"\n>>> Успішно: {ok}, провал: {fail}")
    if fail:
        print("\n=== ПРОВАЛИ (для fix-цикла) ===")
        for num, mn, err in failures:
            print(f"  {num} ({mn:%m.%Y}): {err}")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
