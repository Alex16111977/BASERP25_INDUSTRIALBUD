# -*- coding: utf-8 -*-
# Перевірка що в регістрі БДДС Казни є окремі реквізити СуммаНачисления/СуммаФОТ/СуммаНалогов.
import win32com.client, sys, pywintypes, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
kazna = v8.Connect('Srvr="SQLSERVER";Ref="BuhKazn";Usr="cfo";Pwd="2442"')
S = kazna.String


def com_date(y, m, d, hh=0, mm=0, ss=0):
    return pywintypes.Time(datetime.datetime(y, m, d, hh, mm, ss))


def parse_num(s):
    return float(S(s).replace('\xa0', '').replace(' ', '').replace(',', '.') or "0")


# Метадані БДДС
print("=" * 70)
print("Реквізити регістру БДДС у Казні")
print("=" * 70)
mo = kazna.Метаданные.РегистрыНакопления.Найти("БДДС")
if mo is None:
    print("НЕ ЗНАЙДЕНО")
    sys.exit(1)

print(f"\nИзмерения ({mo.Измерения.Количество()}):")
for i in range(mo.Измерения.Количество()):
    d = mo.Измерения.Получить(i)
    print(f"  {S(d.Имя)}: {S(d.Тип)[:60]}")

print(f"\nРесурсы ({mo.Ресурсы.Количество()}):")
for i in range(mo.Ресурсы.Количество()):
    d = mo.Ресурсы.Получить(i)
    print(f"  {S(d.Имя)}: {S(d.Тип)[:60]}")

print(f"\nРеквизиты ({mo.Реквизиты.Количество()}):")
for i in range(mo.Реквизиты.Количество()):
    d = mo.Реквизиты.Получить(i)
    print(f"  {S(d.Имя)}: {S(d.Тип)[:60]}")

# Спроба запиту з новими колонками
print()
print("=" * 70)
print("Запит з готовими колонками Сумма/СуммаНачисления/СуммаФОТ/СуммаНалогов")
print("=" * 70)

НачПер = com_date(2025, 12, 1)
КонПер = com_date(2025, 12, 31, 23, 59, 59)

q = kazna.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    БДДС.Сотрудник.Наименование КАК СотрудникФИО,
    СУММА(БДДС.Сумма) КАК Сумма,
    СУММА(БДДС.СуммаНачисления) КАК СуммаНачисления,
    СУММА(БДДС.СуммаФОТ) КАК СуммаФОТ,
    СУММА(БДДС.СуммаНалогов) КАК СуммаНалогов
ИЗ РегистрНакопления.БДДС КАК БДДС
ГДЕ (БДДС.Регистратор ССЫЛКА Документ.РаспределениеЗаработнойПлаты
        ИЛИ БДДС.Регистратор ССЫЛКА Документ.РаспределениеФ2)
    И БДДС.Регистратор.КассаДатаС >= &Н
    И БДДС.Регистратор.КассаДатаПо <= &К
СГРУППИРОВАТЬ ПО БДДС.Сотрудник.Наименование
УПОРЯДОЧИТЬ ПО СУММА(БДДС.Сумма) УБЫВ
"""
q.SetParameter("Н", НачПер)
q.SetParameter("К", КонПер)
try:
    r = q.Execute().Выгрузить()
    print(f"  Рядків: {r.Количество()}")

    total_s = 0.0
    total_n = 0.0
    total_f = 0.0
    total_u = 0.0

    for i in range(r.Количество()):
        row = r.Получить(i)
        s = parse_num(row.Сумма)
        n = parse_num(row.СуммаНачисления)
        f = parse_num(row.СуммаФОТ)
        u = parse_num(row.СуммаНалогов)
        total_s += s
        total_n += n
        total_f += f
        total_u += u

    print(f"\n  Тотали:")
    print(f"    Сумма:           {total_s:>15,.2f}")
    print(f"    СуммаНачисления: {total_n:>15,.2f}")
    print(f"    СуммаФОТ:        {total_f:>15,.2f}")
    print(f"    СуммаНалогов:    {total_u:>15,.2f}")

    print(f"\n  Перші 5:")
    for i in range(min(5, r.Количество())):
        row = r.Получить(i)
        print(f"    {S(row.СотрудникФИО)[:30]:30s}  Сум={parse_num(row.Сумма):>11.2f} Нач={parse_num(row.СуммаНачисления):>11.2f} ФОТ={parse_num(row.СуммаФОТ):>11.2f} Нал={parse_num(row.СуммаНалогов):>11.2f}")
except Exception as e:
    print(f"  ERR: {str(e)[:500]}")
