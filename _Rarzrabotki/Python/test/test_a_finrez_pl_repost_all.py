"""Repost усіх документів А_ФинРез_PL після видалення ВключатьДочерние.

Знаходить усі непомічені на видалення документи (проведені і непроведені)
і виконує Записать у режимі Проведення.
Очікувано: 4/4 OK.
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
        ВЫБРАТЬ Ссылка КАК Ref, Номер КАК Num, Дата КАК Dt, Проведен КАК Posted
        ИЗ Документ.А_ФинРез_PL
        ГДЕ НЕ ПометкаУдаления
        УПОРЯДОЧИТЬ ПО Дата
    """
    rows = q.Execute().Выгрузить()
    total = rows.Количество()
    print(f"Знайдено документів (без помітки на видалення): {total}")

    ok, fail = 0, 0
    for i in range(total):
        row = rows.Получить(i)
        ref = row.Ref
        num = row.Num
        dt = row.Dt
        was_posted = row.Posted
        try:
            obj = ref.ПолучитьОбъект()
            if was_posted:
                obj.Записать(erp.РежимЗаписиДокумента.ОтменаПроведения)
                obj = ref.ПолучитьОбъект()
            obj.Записать(erp.РежимЗаписиДокумента.Проведение)
            print(f">>> Док {num} | {dt:%d.%m.%Y} | was_posted={was_posted} - Posted OK")
            ok += 1
        except Exception as e:
            err = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
            print(f">>> Док {num} | {dt:%d.%m.%Y} | was_posted={was_posted} - FAIL: {err}")
            fail += 1

    print(f"\n>>> Успішно: {ok}, провал: {fail}")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
