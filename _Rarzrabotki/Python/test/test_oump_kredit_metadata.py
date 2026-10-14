# -*- coding: utf-8 -*-
"""ОУМП, кредитные договоры: метаданные загружены в BaseERP (перечисление, реквизиты, колонки ТЧ Деньги)."""
import win32com.client, sys
sys.stdout.reconfigure(encoding='utf-8')
erp = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
md = erp.Метаданные
ok = True


def check(cond, msg):
    global ok
    print(("OK  " if cond else "FAIL") + " " + msg)
    ok = ok and bool(cond)


en = md.Перечисления.Найти("А_ВидыКредитныхДоговоров")
check(en is not None, "перечисление А_ВидыКредитныхДоговоров")
if en is not None:
    names = [v.Имя for v in en.ЗначенияПеречисления]
    check(names == ["ТелоКредита", "Проценты", "НераспределеннаяПрибыль"], f"значения {names}")
check(md.Справочники.ДоговорыКонтрагентов.Реквизиты.Найти("А_ВидКредитногоДоговора") is not None, "Договор.А_ВидКредитногоДоговора")
check(md.Справочники.СтатьиДвиженияДенежныхСредств.Реквизиты.Найти("А_ВидКредитногоДоговора") is not None, "СтатьяДДС.А_ВидКредитногоДоговора")
ts = md.Документы.А_ОказаниеУслугМеждуПодразделениями.ТабличныеЧасти.Деньги.Реквизиты
for n in ["КонтрагентРасход", "КонтрагентПриход", "ВидДоговораРасход", "ВидДоговораПриход"]:
    check(ts.Найти(n) is not None, f"Деньги.{n}")
sys.exit(0 if ok else 1)
