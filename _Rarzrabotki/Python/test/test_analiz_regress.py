# -*- coding: utf-8 -*-
# Регресс: «Анализ документов» не падает с новым кодом (ОпределитьСтатьюДисбаланса+НайтиИлиСоздать)
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
ERF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\А_ОбработкаДисбалансаПоПодразделениям.erf"
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
try:
    obr = erp.ВнешниеОтчеты.Создать(ERF, False)
except Exception:
    obr = erp.ВнешниеОтчеты.Создать(ERF)
obr.НачалоПериода = datetime.datetime(2025, 12, 15, 12, 0, 0)
obr.ОкончаниеПериода = datetime.datetime(2025, 12, 15, 12, 0, 0)
obr.ПоказыватьВсе = True
try:
    obr.АнализДокументов()
    print("АнализДокументов выполнен БЕЗ ОШИБОК, строк ТЧ:", obr.ДокументыДисбаланса.Количество())
    print("=== REGRESS PASS ===")
except Exception as e:
    info = getattr(e, 'excepinfo', None)
    print("REGRESS FAIL:", info[2] if info else e)
    print("=== REGRESS FAIL ===")
