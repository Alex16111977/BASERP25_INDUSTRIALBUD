# -*- coding: utf-8 -*-
# Общие хелперы для тестов замены "Неизвестного партнёра"
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'


def connect():
    v8 = win32com.client.Dispatch("V83.COMConnector")
    return v8.Connect(CONN)


def neizv(erp):
    return erp.Справочники.Партнеры.НайтиПоНаименованию("Неизвестный партнер", True)


def uid(erp, ref):
    return erp.String(ref.Ссылка.УникальныйИдентификатор())


def same(erp, a, b):
    """Сравнение ссылок 1С по UUID (Python == сравнивает COM-обёртки, не значения!)."""
    if not erp.ЗначениеЗаполнено(a) or not erp.ЗначениеЗаполнено(b):
        return erp.ЗначениеЗаполнено(a) == erp.ЗначениеЗаполнено(b)
    try:
        return uid(erp, a) == uid(erp, b)
    except Exception:
        return False


def count_refs(erp, partner):
    arr = erp.NewObject("Массив")
    arr.Добавить(partner)
    res = erp.НайтиПоСсылкам(arr)
    agg = {}
    for i in range(res.Количество()):
        try:
            fn = res.Получить(i).Метаданные.ПолноеИмя()
        except Exception:
            fn = "?"
        agg[fn] = agg.get(fn, 0) + 1
    return res.Количество(), agg


def balance_snapshot(erp):
    """Σ остатков расчётов — для сверки до/после."""
    q = erp.NewObject("Запрос")
    q.Text = (
        'ВЫБРАТЬ "РСК" КАК Р, СУММА(СуммаОстаток) КАК С '
        'ИЗ РегистрНакопления.РасчетыСКлиентами.Остатки '
        'ОБЪЕДИНИТЬ ВСЕ '
        'ВЫБРАТЬ "РСП", СУММА(СуммаОстаток) '
        'ИЗ РегистрНакопления.РасчетыСПоставщиками.Остатки'
    )
    out = {}
    s = q.Execute().Выбрать()
    while s.Следующий():
        out[s.Р] = float(s.С or 0)
    return out


if __name__ == "__main__":
    e = connect()
    p = neizv(e)
    print("UID", uid(e, p))
    total, agg = count_refs(e, p)
    print("refs всего:", total)
    print("bal:", balance_snapshot(e))
