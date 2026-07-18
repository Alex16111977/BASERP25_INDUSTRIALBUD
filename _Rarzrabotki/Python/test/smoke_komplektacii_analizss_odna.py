# -*- coding: utf-8 -*-
# Фаза 14 smoke: 5-й звіт «Аналіз СС (одна одиниця)» — ОДНА колонка «Одиниця виміру»
# з пріоритетом ЕдиницаСС -> Единица. Спец 000000004 / «СТІЛ МД МХП ОРІЛЬ виробничий» / 07.07.2026.
import win32com.client, sys, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
EPF = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Создания комплектаций для бух учета.epf"
обр = buh.ВнешниеОбработки.Создать(EPF, False)
обр.Период = datetime.datetime(2026, 7, 7)
обр.Спецификация = buh.Справочники.СтруктураСебестоимости.НайтиПоКоду("000000004").Ссылка
выб = buh.Справочники.Склады.Выбрать()
while выб.Следующий():
    н = выб.Наименование
    if "СТІЛ" in н and "МХП" in н and "ОР" in н and "виробнич" in н:
        обр.СкладыОстатков.Добавить().Склад = выб.Ссылка
обр.РассчитатьАнализ()
тч = обр.ТабличнаяЧастьОстатков

# очікувана коалесцентна одиниця по кожній номенклатурі (пріоритет ЕдиницаСС -> Единица)
expected = {}
subst = {}   # субститути: ЕдиницаСС порожня, Единица заповнена
norm = {}    # норм-картки: ЕдиницаСС заповнена
for i in range(тч.Количество()):
    с = тч.Получить(i)
    имя = с.Номенклатура.Наименование.strip()
    есть_сс = not с.ЕдиницаСС.Пустая()
    есть_ед = not с.Единица.Пустая()
    if есть_сс:
        exp = с.ЕдиницаСС.Наименование.strip()
        norm[имя] = exp
    elif есть_ед:
        exp = с.Единица.Наименование.strip()
        subst[имя] = exp
    else:
        exp = ""
    expected[имя] = exp
print(f"ТЧ: рядків={тч.Количество()}; норм-карток(ЕдиницаСС)={len(norm)}; субститутів(тільки Единица)={len(subst)}")

табдок = обр.СформироватьПечатьАнализССОдна(тч)
h, w = табдок.ВысотаТаблицы, табдок.ШиринаТаблицы
print(f"ВысотаТаблицы={h} ШиринаТаблицы={w}")

# знайти рядок-шапку заголовків і колонку «Одиниця виміру»
hdr_row = None; col_unit = None; col_zal = None; col_ss = None
for r in range(1, min(8, h) + 1):
    for c in range(1, w + 1):
        t = str(табдок.Область(r, c, r, c).Текст).strip()
        if t == "Одиниця виміру":
            hdr_row, col_unit = r, c
        if t == "Залишок":
            col_zal = c
        if t == "Згідно з СС":
            col_ss = c
assert col_unit is not None, "немає колонки «Одиниця виміру»"
# не має бути двох окремих одиниць
allcells = " | ".join(str(табдок.Область(r, c, r, c).Текст) for r in range(1, min(8, h)+1) for c in range(1, w+1))
assert "Одиниця виміру залишків" not in allcells and "Одиниця виміру СС" not in allcells, "є подвійні одиниці — має бути одна"
print(f"Колонка «Одиниця виміру» c={col_unit}; Залишок c={col_zal}; Згідно з СС c={col_ss}")
assert col_zal < col_unit < col_ss, f"порядок невірний: Залишок={col_zal} < Одиниця={col_unit} < ЗгідноСС={col_ss}"

# коалесценція end-to-end: для рядків-номенклатур (col1==ном) значення unit-колонки == expected
checked = 0; checked_norm = 0; checked_subst = 0; bad = []
for r in range(hdr_row + 1, h + 1):
    имя = str(табдок.Область(r, 1, r, 1).Текст).strip()
    if имя in expected and expected[имя]:
        got = str(табдок.Область(r, col_unit, r, col_unit).Текст).strip()
        if got != expected[имя]:
            bad.append((имя, got, expected[имя]))
        else:
            checked += 1
            if имя in norm: checked_norm += 1
            if имя in subst: checked_subst += 1
assert not bad, f"коалесценція невірна (перші 5): {bad[:5]}"
assert checked_norm > 0, "не перевірено жодної норм-картки (ЕдиницаСС-гілка)"
print(f"Коалесценція OK: перевірено рядків={checked} (норм-гілка={checked_norm}, субститут-гілка={checked_subst})")
if checked_subst == 0:
    print("  (субститутів у цьому наборі немає — гілка Единица перевірена лише логікою ТЧ)")

# шапка не янтарна
AMBER = (255, 230, 153); ok_head = False
for r in range(1, min(8, h) + 1):
    for c in range(1, w + 1):
        обл = табдок.Область(r, c, r, c)
        if "Аналіз залишків" in str(обл.Текст):
            цф = обл.ЦветФона
            assert (цф.R, цф.G, цф.B) != AMBER, f"шапка янтарна {(цф.R,цф.G,цф.B)}"
            ok_head = True; break
    if ok_head: break
assert ok_head, "не знайдено заголовок звіту"
print("PASS: одна колонка «Одиниця виміру», коалесценція ЕдиницаСС->Единица, порядок, шапка без янтаря")
