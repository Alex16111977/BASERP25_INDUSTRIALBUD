# -*- coding: utf-8 -*-
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

def tch_fields(doc, tch):
    md = erp.Метаданные.Документы.Найти(doc)
    for t in md.ТабличныеЧасти:
        if t.Имя == tch:
            return [r.Имя for r in t.Реквизиты]
    return None

print("ВводОстатков.РасчетыСПартнерами:")
print(" ", tch_fields("ВводОстатков", "РасчетыСПартнерами"))
print("ВводОстатков.БанковскиеСчета:")
print(" ", tch_fields("ВводОстатков", "БанковскиеСчета"))
print("ВводОстатковВнеоборотныхАктивов.ОС (первые 12):")
print(" ", tch_fields("ВводОстатковВнеоборотныхАктивов", "ОС")[:12])
print("НачальнаяЗадолженностьПоЗарплате.ЗадолженностьПоЗарплате:")
print(" ", tch_fields("НачальнаяЗадолженностьПоЗарплате", "ЗадолженностьПоЗарплате"))

# Проверим, что регистр СоответствияОбъектовИнформационныхБаз есть (для UUID-маппинга)
ir = erp.Метаданные.РегистрыСведений.Найти("СоответствияОбъектовИнформационныхБаз")
print("РС СоответствияОбъектовИнформационныхБаз:", "OK" if ir is not None else "NO")
