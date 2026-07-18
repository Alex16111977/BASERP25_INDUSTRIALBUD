# -*- coding: utf-8 -*-
# Шукаємо як пов'язати Удержания -> ИсполнительныйЛист в ZUP_1
import win32com.client, sys, pywintypes, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
zup = v8.Connect('Srvr="localhost";Ref="zup";Usr="cfo";Pwd="2442"')
S = zup.String

def com_date(y, m, d):
    return pywintypes.Time(datetime.datetime(y, m, d))

print("=" * 80)
print("КРОК 1: всі реквізити і виміри РегистрРасчета.УдержанияРаботниковОрганизаций")
print("=" * 80)
try:
    mo = zup.Метаданные.НайтиПоПолномуИмени("РегистрРасчета.УдержанияРаботниковОрганизаций")
    if mo is None:
        print("  НЕ ЗНАЙДЕНО")
    else:
        print(f"  Регістр: {S(mo.Имя)}")
        print(f"\n  Измерения ({mo.Измерения.Количество()}):")
        for i in range(mo.Измерения.Количество()):
            d = mo.Измерения.Получить(i)
            print(f"    {S(d.Имя)}: {S(d.Тип)[:80]}")
        print(f"\n  Ресурсы ({mo.Ресурсы.Количество()}):")
        for i in range(mo.Ресурсы.Количество()):
            d = mo.Ресурсы.Получить(i)
            print(f"    {S(d.Имя)}: {S(d.Тип)[:80]}")
        print(f"\n  Реквизиты ({mo.Реквизиты.Количество()}):")
        for i in range(mo.Реквизиты.Количество()):
            d = mo.Реквизиты.Получить(i)
            print(f"    {S(d.Имя)}: {S(d.Тип)[:80]}")
except Exception as e:
    print(f"  ERROR: {str(e)[:300]}")

print()
print("=" * 80)
print("КРОК 2: пошук інформаційних регістрів з 'Удержан' або 'Исполн'")
print("=" * 80)
try:
    reg_sved = zup.Метаданные.РегистрыСведений
    cnt = reg_sved.Количество()
    print(f"  Всього регістрів відомостей: {cnt}")
    found = []
    for i in range(cnt):
        r = reg_sved.Получить(i)
        nm = S(r.Имя)
        if ("Удерж" in nm or "Исполн" in nm or "Плановы" in nm.lower()):
            found.append(nm)
    print(f"  З 'Удерж'/'Исполн'/'Плановы' у назві ({len(found)}):")
    for nm in found:
        print(f"    {nm}")
except Exception as e:
    print(f"  ERROR: {str(e)[:300]}")

print()
print("=" * 80)
print("КРОК 3: реальний запит до РегистрРасчета.УдержанияРаботниковОрганизаций")
print("=" * 80)
try:
    q = zup.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ ПЕРВЫЕ 5 *
    ИЗ РегистрРасчета.УдержанияРаботниковОрганизаций КАК Удержания
    ГДЕ Удержания.ПериодРегистрации МЕЖДУ &Н И &К
        И Удержания.Результат <> 0
    """
    q.SetParameter("Н", com_date(2025, 12, 1))
    q.SetParameter("К", com_date(2025, 12, 31))
    r = q.Execute().Выгрузить()
    print(f"  Рядків: {r.Количество()}")
    if r.Количество() > 0:
        print(f"  Колонки:")
        for i in range(r.Колонки.Количество()):
            col = r.Колонки.Получить(i)
            print(f"    {S(col.Имя)}: {S(col.ТипЗначения)[:80]}")
        print()
        print(f"  Перший рядок:")
        row = r.Получить(0)
        for i in range(r.Колонки.Количество()):
            col = r.Колонки.Получить(i)
            val = getattr(row, S(col.Имя))
            print(f"    {S(col.Имя)}: {S(val)[:100]}")
except Exception as e:
    print(f"  ERROR: {str(e)[:500]}")
