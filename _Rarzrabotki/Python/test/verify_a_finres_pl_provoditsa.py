# -*- coding: utf-8 -*-
"""Verify: А_ФинРез_PL проводиться без помилки після фіксу line 81."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client, pythoncom

pythoncom.CoInitialize()
conn = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String

# Берем тестовий документ № 00000000003 від 31.03.2026 (стуб для Task 7 з UI)
q = conn.NewObject("Запрос")
q.Текст = """ВЫБРАТЬ ПЕРВЫЕ 5 Б.Ссылка КАК Б, Б.Номер, Б.Дата, Б.Проведен, Б.ПометкаУдаления
ИЗ Документ.А_ФинРез_PL КАК Б
УПОРЯДОЧИТЬ ПО Б.Дата"""
r = q.Выполнить().Выгрузить()
print(f"[INFO] Знайдено {r.Количество()} документів А_ФинРез_PL:")
PROVED = conn.PredefinedValue("РежимЗаписиДокумента.Проведение")

for i in range(r.Количество()):
    rec = r.Получить(i)
    print(f"  №{rec.Номер.strip()} від {rec.Дата.strftime('%Y-%m-%d')} Проведен={rec.Проведен} ПометкаУдаления={rec.ПометкаУдаления}")
    if rec.ПометкаУдаления:
        print(f"    (skip — помічений)")
        continue
    try:
        rec.Б.ПолучитьОбъект().Записать(PROVED)
        print(f"    ✅ Перепроведено успішно")
    except Exception as e:
        err = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
        print(f"    ❌ ПОМИЛКА: {err}")
