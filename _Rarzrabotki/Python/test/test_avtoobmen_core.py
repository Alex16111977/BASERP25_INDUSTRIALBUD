# -*- coding: utf-8 -*-
"""Rule #-1: headless-обмен по узлу через серверный метод БСП (без окна помощника).
Проверяем на BaseERP, что ОбменДаннымиСервер.ВыполнитьОбменДаннымиДляУзлаИнформационнойБазы
отрабатывает без исключения и продвигает регистр состояний успешных обменов."""
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')


def uzel(plan, kod):
    menedzher = getattr(erp.ПланыОбмена, plan)
    return menedzher.НайтиПоКоду(kod)


def poslednij_uspeh(node):
    q = erp.NewObject("Запрос")
    q.Text = ("ВЫБРАТЬ МАКСИМУМ(С.ДатаОкончания) КАК Д "
              "ИЗ РегистрСведений.СостоянияУспешныхОбменовДанными КАК С "
              "ГДЕ С.УзелИнформационнойБазы = &У")
    q.УстановитьПараметр("У", node)
    r = q.Выполнить().Выбрать()
    return r.Д if r.Следующий() else None


def obmen(node):
    P = erp.ОбменДаннымиСервер.ПараметрыОбмена()
    P.Вставить("ВидТранспортаСообщенийОбмена", erp.Перечисления.ВидыТранспортаСообщенийОбмена.FILE)
    P.Вставить("ВыполнятьВыгрузку", True)
    P.Вставить("ВыполнятьЗагрузку", True)
    # Отказ — out-параметр по ссылке; win32com вернёт его в результате
    res = erp.ОбменДаннымиСервер.ВыполнитьОбменДаннымиДляУзлаИнформационнойБазы(node, P, False)
    return res


# --- только bas_industrialbud на первом прогоне ---
n = uzel("ОбменУправлениеПредприятиемБухгалтерия20", "БП")
print("Узел:", n.Наименование)
print("Последний успешный обмен ДО:", poslednij_uspeh(n))
try:
    res = obmen(n)
    print("Вызов выполнен без исключения. Возврат (Отказ):", res)
except Exception as e:
    info = e.excepinfo[2] if getattr(e, 'excepinfo', None) else e
    print("FAIL:", info)
    sys.exit(1)
print("Последний успешный обмен ПОСЛЕ:", poslednij_uspeh(n))
print("OK")
