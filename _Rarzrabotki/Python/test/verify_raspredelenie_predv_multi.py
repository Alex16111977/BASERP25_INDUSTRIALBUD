# -*- coding: utf-8 -*-
"""
Rule #-1 БЕЗ записи: предварительный режим РаспределитьРасходыНаФинансовыйРезультат на нескольких
документах. Проверяем:
  - Σ-инвариант: ИТОГО РезультатРаспределения == Σ текущих движений документа (оригинальный результат);
  - НЕТ строки (пусто) с ненулевой суммой (значит нет двойного счёта / нетит);
  - разрез по форме (Ф1/Ф2 для ЗП-статей; пусто→Форма1 для статей без формленного источника).
"""
import sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
import win32com.client, pythoncom
from win32com.client import VARIANT

erp = win32com.client.Dispatch("V83.COMConnector").Connect(
    'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
TOL = 0.01

# (Номер, год, месяц, день) — тест-документы РаспределениеПрочихЗатрат
DOCS = [
    ("00000000279", 2025, 12, 31),  # Зарплата управ — Ф1+Ф2 (эталон)
    ("00000000277", 2025, 12, 31),  # Зарплата управ
    ("00000000321", 2025, 12, 31),  # Зарплата ИТР
    ("00000001225", 2026, 3, 31),   # Зарплата произв (март)
    ("00000000012", 2025, 12, 31),  # РКО_Плата (банк, источник без формы → ожид. пусто)
]


def forma(f):
    if not erp.ЗначениеЗаполнено(f):
        return "(пусто)"
    x = erp.XMLСтрока(f)
    return "Форма1" if "Форма1" in x else ("Форма2" if "Форма2" in x else x)


def get_doc(num, y, m, d):
    q = erp.NewObject("Запрос")
    q.Текст = ('ВЫБРАТЬ Р.Ссылка КАК С, Р.Организация КАК Орг ИЗ Документ.РаспределениеПрочихЗатрат КАК Р '
               'ГДЕ Р.Номер="%s" И Р.Дата=ДАТАВРЕМЯ(%d,%d,%d,23,59,59)' % (num, y, m, d))
    s = q.Выполнить().Выбрать()
    if not s.Следующий():
        return None, None
    return s.С, s.Орг


def current_sum(ref):
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Рег", ref)
    q.Текст = ("ВЫБРАТЬ СУММА(П.Сумма) КАК С ИЗ РегистрНакопления.ПрочиеРасходы КАК П "
               "ГДЕ П.Регистратор = &Рег И П.Активность")
    r = q.Выполнить().Выбрать(); r.Следующий()
    return float(r.С) if r.С is not None else 0.0


def result_by_form(tm):
    q = erp.NewObject("Запрос")
    q.МенеджерВременныхТаблиц = tm
    q.Текст = ("ВЫБРАТЬ Т.ФормаPL КАК ФормаPL, СУММА(Т.Сумма) КАК С ИЗ РезультатРаспределения КАК Т "
               "СГРУППИРОВАТЬ ПО Т.ФормаPL")
    t = q.Выполнить().Выгрузить()
    res = {}
    for i in range(t.Количество()):
        r = t.Получить(i)
        res[forma(r.ФормаPL)] = float(r.С) if r.С is not None else 0.0
    return res


print("%-14s %14s %14s %14s %14s  %s" % ("Документ", "ТекущаяΣ", "РезультатΣ", "Ф1", "Ф2", "Вердикт"))
for num, y, m, d in DOCS:
    ref, org = get_doc(num, y, m, d)
    if ref is None:
        print("%-14s  НЕ НАЙДЕН" % num)
        continue
    cur = current_sum(ref)
    орги = erp.NewObject("Array"); орги.Добавить(org)
    mvt_var = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None)
    try:
        erp.Документы.РаспределениеПрочихЗатрат.РаспределитьРасходыНаФинансовыйРезультат(
            datetime.datetime(y, m, d), орги, mvt_var, ref)
        res = result_by_form(mvt_var.value)
    except Exception as e:
        ei = getattr(e, 'excepinfo', None)
        print("%-14s  ОШИБКА: %s" % (num, ei[2] if ei else e))
        continue
    f1 = res.get("Форма1", 0.0); f2 = res.get("Форма2", 0.0); pu = res.get("(пусто)", 0.0)
    tot = f1 + f2 + pu
    sigma_ok = abs(tot - cur) <= TOL          # Σ-инвариант
    pusto_ok = abs(pu) <= TOL                  # нет ненулевого (пусто) → нет двойного счёта
    verdict = "OK" if (sigma_ok and pusto_ok) else ("FAIL Σ" if not sigma_ok else "FAIL пусто=%.2f" % pu)
    print("%-14s %14.2f %14.2f %14.2f %14.2f  %s" % (num, cur, tot, f1, f2, verdict))
