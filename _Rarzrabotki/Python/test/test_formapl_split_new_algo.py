# -*- coding: utf-8 -*-
"""
RED-модель нового алгоритма разметки ФормаPL (Часть A промта
PROMPT_ФормаPL_split_и_ОрганизацияБух.md).

Читает реальные данные документа А_ОтражениеЗПпоКазне №000000001 (сохранённые ТЧ),
применяет НОВЫЙ алгоритм для Постернака и Шишкиной и сверяет с эталонами §4.

Алгоритм (для ФЛ с НачисленияУпр.Сумма ≠ 0):
  Ф1 = max(0, min(NET, GROSS_бух − СуммаНДФЛ))   # официальная take-home
  Ф2 = NET − Ф1                                  # управленческий остаток
  «Зарплата управл» Ф1 → по коэф Форма1-начислений Казны (ОргБух из Казны)
  «Зарплата управл» Ф2 → премии (А_Начисление=Истина) как есть + остаток по коэф не-премий (ОргБух пусто)
  «Удержание» = СуммаНДФЛ → целиком Ф1 по подразделениям
"""
import sys
from decimal import Decimal, ROUND_HALF_UP
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

def r2(x):
    return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def raspredelit(total, items):
    """items = [(key, weight), ...]; метод «остаток на последнем» как в BSL.
    Возвращает {key: сумма}."""
    res = {}
    s = sum(w for _, w in items)
    if s == 0 or not items:
        return res
    ost = total
    for i, (k, w) in enumerate(items):
        if i == len(items) - 1:
            part = r2(ost)
        else:
            part = r2(total * w / s)
            ost = r2(ost - part)
        res[k] = res.get(k, 0) + part
    return res

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

def q(text):
    z = erp.NewObject("Запрос")
    z.Текст = text
    return z.Выполнить().Выгрузить()

def scalar(text):
    t = q(text)
    return float(t[0][0]) if t.Количество() > 0 and t[0][0] is not None else 0.0

DOC = 'Ф.Ссылка.Номер = "000000001" И Ф.Ссылка.Дата = ДАТАВРЕМЯ(2025,12,31,12,0,0)'

def analyze(fio):
    net = scalar(f'ВЫБРАТЬ СУММА(Ф.Сумма) ИЗ Документ.А_ОтражениеЗПпоКазне.НачисленияУпр КАК Ф '
                 f'ГДЕ {DOC} И Ф.ФизЛицо.Наименование ПОДОБНО "{fio}%"')
    gross = scalar(f'ВЫБРАТЬ СУММА(Ф.Сумма) ИЗ Документ.А_ОтражениеЗПпоКазне.НачисленияБух КАК Ф '
                   f'ГДЕ {DOC} И Ф.ФизЛицо.Наименование ПОДОБНО "{fio}%"')
    ndfl = scalar(f'ВЫБРАТЬ СУММА(Ф.СуммаНДФЛ) ИЗ Документ.А_ОтражениеЗПпоКазне.НалогиБухгалтерия КАК Ф '
                  f'ГДЕ {DOC} И Ф.Сотрудник.ФизическоеЛицо.Наименование ПОДОБНО "{fio}%" '
                  f'И (Ф.ТипНалога = ЗНАЧЕНИЕ(Перечисление.ТипыНалогов.НДФЛ) '
                  f'ИЛИ Ф.ТипНалога = ЗНАЧЕНИЕ(Перечисление.ТипыНалогов.ВоенныйСбор))')

    # Казна-строки с признаком премии и удержания
    kazna = q(
        f'ВЫБРАТЬ ПРЕДСТАВЛЕНИЕ(Ф.ФормаPL) КАК Форма, ПРЕДСТАВЛЕНИЕ(Ф.Подразделение) КАК Подр, '
        f'ПРЕДСТАВЛЕНИЕ(Ф.Организация) КАК Орг, Ф.СуммаНачисления КАК СН, Ф.СуммаНалогов КАК СНал, '
        f'Ф.СтатьяДвиженияДенежныхСредств.А_СпособОтраженияЗарплатыВБухучете.А_ЭтоУдержание КАК Удерж, '
        f'ЕСТЬNULL(ВЫРАЗИТЬ(Ф.ДокРаспределениеЗП КАК Документ.РаспределениеФ2).А_Начисление, ЛОЖЬ) КАК Премия '
        f'ИЗ Документ.А_ОтражениеЗПпоКазне.РаспределениеКазна КАК Ф '
        f'ГДЕ {DOC} И Ф.Сотрудник.ФизическоеЛицо.Наименование ПОДОБНО "{fio}%"')

    ф1_нач, ф2_прем, ф2_неп, удерж_подр = [], {}, [], {}
    for r in kazna:
        forma, podr, org = S(r.Форма), S(r.Подр), S(r.Орг)
        sn, snal, udr, prem = float(r.СН), float(r.СНал), bool(r.Удерж), bool(r.Премия)
        if forma == "Форма1" and not udr and sn != 0:
            ф1_нач.append((podr, org, sn))
        elif forma == "Форма1" and udr:
            удерж_подр[podr] = удерж_подр.get(podr, 0) + snal
        elif forma == "Форма2" and sn != 0:
            if prem:
                ф2_прем[podr] = ф2_прем.get(podr, 0) + sn
            else:
                ф2_неп.append((podr, sn))

    ф1_total = max(0.0, min(net, r2(gross - ndfl)))
    ф2_total = r2(net - ф1_total)

    print(f"\n=== {fio} ===")
    print(f"  NET={net:,.2f}  GROSS={gross:,.2f}  НДФЛ={ndfl:,.2f}")
    print(f"  Ф1total={ф1_total:,.2f}  Ф2total={ф2_total:,.2f}")

    # Ф1 «Зарплата управл» по коэф Форма1-начислений; ОргБух = орг строки
    орг_по_подр = {}
    for podr, org, sn in ф1_нач:
        орг_по_подр[podr] = org
    ф1_распр = raspredelit(ф1_total, [(podr, sn) for podr, _, sn in ф1_нач])

    # Ф2 «Зарплата управл»: премии как есть + остаток по коэф не-премий
    sum_prem = sum(ф2_прем.values())
    ost = r2(ф2_total - sum_prem)
    ост_распр = raspredelit(ost, [(podr, sn) for podr, sn in ф2_неп])
    ф2_итог = dict(ф2_прем)
    for podr, val in ост_распр.items():
        ф2_итог[podr] = ф2_итог.get(podr, 0) + val

    # Удержание Ф1 = НДФЛ по коэф удержания Казны
    удерж_распр = raspredelit(ndfl, list(удерж_подр.items()))

    print("  Ф1 «Зарплата управл»:", {k: round(v, 2) for k, v in ф1_распр.items()},
          "ОргБух:", орг_по_подр)
    print("  Ф2 «Зарплата управл»:", {k: round(v, 2) for k, v in ф2_итог.items()}, "(ОргБух пусто)")
    print("  Ф1 «Удержание»:", {k: round(v, 2) for k, v in удерж_распр.items()})

    # Проверки инварианта
    s_f1 = sum(ф1_распр.values()); s_f2 = sum(ф2_итог.values())
    inv = r2(s_f1 + s_f2)
    print(f"  ИНВАРИАНТ Ф1+Ф2={inv:,.2f}  (NET={net:,.2f})  ->",
          "OK" if abs(inv - net) < 0.02 else "FAIL")
    f1_all = r2(s_f1 + sum(удерж_распр.values()))
    print(f"  Ф1всего (нач+удерж)={f1_all:,.2f}  (GROSS={gross:,.2f}) ->",
          "OK" if abs(f1_all - gross) < 0.05 else "FAIL")
    return inv, net, f1_all, gross

print("RED-модель нового алгоритма ФормаPL split")
for fio in ["Постернак", "Шишкіна"]:
    analyze(fio)
