# -*- coding: utf-8 -*-
"""СКРИПТ 33 — Сравнение ІБ00-007576 в двух базах: BASERP25 (наша) vs BaseERPRazr (типовая).

Цель: понять разницу в РСППС/ПАП движениях для одного и того же документа в двух базах с
идентичным кодом проведения.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from datetime import datetime, timedelta
import win32com.client, pythoncom

pythoncom.CoInitialize()
v8 = win32com.client.Dispatch("V83.COMConnector")

print("=" * 110)
print("СКРИПТ 33 — Сравнение ІБ00-007576 в BASERP25 vs BaseERPRazr")
print("=" * 110)

def analyze(label, conn_str):
    print(f"\n{'-' * 110}\n[{label}]\n{'-' * 110}")
    try:
        erp = v8.Connect(conn_str)
    except Exception as e:
        print(f"  ОШИБКА подключения: {str(e)[:200]}"); return None
    S = erp.String

    dt = datetime(2025, 12, 19, 18, 11, 56)
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Ном", "ІБ00-007576")
    q.УстановитьПараметр("Н1", dt - timedelta(days=1))
    q.УстановитьПараметр("Н2", dt + timedelta(days=1))
    q.Текст = """ВЫБРАТЬ Ссылка, Номер, Дата, Проведен, ПометкаУдаления, Партнер,
        Подразделение, СуммаДокумента, ОбъектРасчетовУпр, Договор,
        ПорядокРасчетов
        ИЗ Документ.ПриобретениеТоваровУслуг
        ГДЕ Номер = &Ном И Дата МЕЖДУ &Н1 И &Н2"""
    r = q.Выполнить().Выгрузить()
    if r.Количество() == 0:
        print("  Документ НЕ найден"); return None
    rec = r.Получить(0)
    ref = rec.Ссылка
    print(f"  Ссылка: {S(ref)}")
    print(f"  Проведен:    {bool(rec.Проведен)}, Помечен: {bool(rec.ПометкаУдаления)}")
    print(f"  Партнер:     {S(rec.Партнер)}")
    print(f"  Подразделение:{S(rec.Подразделение)}")
    print(f"  Сумма:       {rec.СуммаДокумента}")
    print(f"  Договор:     {S(rec.Договор)}")
    print(f"  ПорядокРасч: {S(rec.ПорядокРасчетов)}")
    print(f"  ОбъектРасчетовУпр: {S(rec.ОбъектРасчетовУпр)}")

    # Движения по регистрам
    def cnt(reg, field="Регистратор"):
        qq = erp.NewObject("Запрос")
        qq.УстановитьПараметр("Д", ref)
        qq.Текст = f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К ИЗ {reg} ГДЕ {field} = &Д"
        try:
            return int(qq.Выполнить().Выгрузить().Получить(0).К or 0)
        except Exception as e:
            return -1

    print(f"\n  Движения:")
    for reg, fld in [
        ("РегистрНакопления.ПрочиеАктивыПассивы", "Регистратор"),
        ("РегистрНакопления.РасчетыСПоставщикамиПоСрокам", "ДокументРегистратор"),
        ("РегистрНакопления.РасчетыСПоставщиками", "Регистратор"),
        ("РегистрНакопления.ТоварыНаСкладах", "Регистратор"),
    ]:
        n = cnt(reg, fld)
        marker = "⚠" if (n == 0 and "ПоСрокам" in reg) else " "
        print(f"  {marker} {reg:<55} {fld:<22} {n}")

    # Детально ТЧ Товары (первая строка)
    тов = ref.ПолучитьОбъект().Товары
    print(f"\n  ТЧ Товары ({тов.Количество()} стр), первая:")
    if тов.Количество() > 0:
        t = тов.Получить(0)
        print(f"    Номенклатура: {S(t.Номенклатура)}")
        print(f"    Сумма:        {t.Сумма}")
        print(f"    ОбъектРасчетов:{S(t.ОбъектРасчетов)}")
        print(f"    СтатьяРасходов:{S(t.СтатьяРасходов)}")
        print(f"    Подразделение:{S(t.Подразделение)}")
    return ref


analyze("BASERP25 (наша)",   'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
analyze("BaseERPRazr (типовая)", 'Srvr="SQLSERVER";Ref="BaseERPRazr";Usr="Администратор";Pwd="24043"')
