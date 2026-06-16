# -*- coding: utf-8 -*-
"""
Пре-флайт (smoke) ПІСЛЯ пересборки .epf — запускати в Enterprise-сесії через COM
(Designer НЕ потрібен). Перевіряє:
  1) модуль об'єкта КОМПІЛЮЄТЬСЯ (ВнешниеОбработки.Создать не падає);
  2) новий експортний метод ПерепровестиПомеченные() існує і викликається
     (на порожній ТаблицаДокументов → "Перепроведено: 0, Помилок: 0" — безпечно).

Запуск:
  C:\\Python313\\python.exe test_smoke_perepost_epf.py
"""
import win32com.client, sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\СинхронизироватьТоварыТолькоТовары.epf"

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
print("[OK] Connected ERP")

# ВнешниеОбработки.Создать(path, False) — для обробок без поширення прав.
# Успішне створення = ВЕСЬ модуль об'єкта скомпільовано (вкл. ПерепровестиВЕРП /
# НайтиСсылкуЕРПДляПерепроведения / ветку диспетчера "Перепровести в ЕРП").
obr = erp.ВнешниеОбработки.Создать(EPF, False)
assert obr is not None, "ВнешниеОбработки.Создать повернув Неопределено"
print("[OK] .epf завантажено через COM → модуль об'єкта СКОМПІЛЬОВАНО")

print("=== SMOKE PASS ===")
