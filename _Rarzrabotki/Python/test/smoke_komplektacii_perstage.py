# -*- coding: utf-8 -*-
# Фаза 16 smoke: коректна по-етапна розшивка (Этап×Еталон) звіту «за етапами».
# Спец 000000005 (є багатоетапні картки). Перевірка:
#  1) підсумки (Итого) звіту-за-етапами == не-етапного «Аналіз СС (одна одиниця)» (розшивка зберігає Σ);
#  2) багатоетапна картка «Профнастил ПС-10» розбита на >=2 этапи.
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')
EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Создания комплектаций для бух учета.epf"
обр = buh.ВнешниеОбработки.Создать(EPF, False)
обр.Период = datetime.datetime(2026, 6, 25)
обр.Спецификация = buh.Справочники.СтруктураСебестоимости.НайтиПоКоду("000000005").Ссылка
выб = buh.Справочники.Склады.Выбрать()
while выб.Следующий():
    н = выб.Наименование
    if "МХП" in н and "ОР" in н and "виробнич" in н:
        обр.СкладыОстатков.Добавить().Склад = выб.Ссылка
обр.РассчитатьАнализ()

def num(s):
    s = str(s).replace("\xa0", "").replace(" ", "").replace(",", ".").strip()
    try:
        return float(s)
    except ValueError:
        return None

def totals(t):
    h, w = t.ВысотаТаблицы, t.ШиринаТаблицы
    colhdr = {}
    for r in range(1, min(9, h) + 1):
        for c in range(1, w + 1):
            txt = str(t.Область(r, c, r, c).Текст).strip()
            for hd in ("Залишок", "Згідно з СС", "В нормі", "Понад норму", "Факт. списання"):
                if txt == hd and hd not in colhdr:
                    colhdr[hd] = c
    # рядок «Итого» (grand total) — знизу
    for r in range(h, 1, -1):
        c1 = str(t.Область(r, 1, r, 1).Текст).strip()
        if "Итог" in c1 or "Разом" in c1:
            return {hd: num(t.Область(r, col, r, col).Текст) for hd, col in colhdr.items()}, r
    return None, None

t_et = обр.СформироватьПечатьПланФактЕтапи(обр.ТабличнаяЧастьОстатков)
t_od = обр.СформироватьПечатьАнализССОдна(обр.ТабличнаяЧастьОстатков)
tot_et, r_et = totals(t_et)
tot_od, r_od = totals(t_od)
print("Итого (за етапами):", tot_et, "рядок", r_et)
print("Итого (одна одиниця):", tot_od, "рядок", r_od)
assert tot_et and tot_od, "не знайдено рядок «Итого» в одному зі звітів"

# 1) підсумки збережено (розшивка не змінює Σ)
for hd in ("Залишок", "Згідно з СС", "В нормі", "Понад норму", "Факт. списання"):
    a, b = tot_et.get(hd), tot_od.get(hd)
    if a is None or b is None:
        continue
    assert abs(a - b) <= 0.5, f"підсумок «{hd}» розійшовся: за-етапами={a} vs одна-одиниця={b}"
print("Підсумки за-етапами == одна-одиниця (розшивка зберігає Σ) — OK")

# 2) багатоетапна картка розбита
cnt = 0
for r in range(1, t_et.ВысотаТаблицы + 1):
    for c in range(1, t_et.ШиринаТаблицы + 1):
        if str(t_et.Область(r, c, r, c).Текст).strip() == "Профнастил ПС-10":
            cnt += 1
print("Профнастил ПС-10 рядків (этапів):", cnt)
assert cnt >= 2, f"багатоетапна картка НЕ розбита по этапах (cnt={cnt})"

print("PASS: по-етапна розшивка (Этап×Еталон) коректна — Σ збережено, багатоетапна картка розбита")
