# -*- coding: utf-8 -*-
"""Rule #-1: пакетный запрос каталога групп/статей для ДополнитьВидКакВExcel v2 (1:1 из BSL)."""
import win32com.client, sys
sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = erp.NewObject("Запрос")
q.Text = """ВЫБРАТЬ Ссылка, Код, Наименование ИЗ Справочник.А_ГруппаСтатей_PL ГДЕ НЕ ПометкаУдаления
;
ВЫБРАТЬ
	Ст.Ссылка КАК Ссылка,
	Ст.Код КАК Код,
	Ст.Наименование КАК Наименование,
	Ст.Группа КАК Группа,
	Ст.Сорт КАК Сорт
ИЗ Справочник.А_Статьи_PL КАК Ст
ГДЕ НЕ Ст.ПометкаУдаления"""
try:
    pak = q.ВыполнитьПакет()
    g = pak.Получить(0).Выгрузить()
    a = pak.Получить(1).Выгрузить()
    print(f"OK: групп={g.Количество()} статей={a.Количество()}")
except Exception as e:
    msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
    print("FAIL:", msg)
    sys.exit(1)
