# -*- coding: utf-8 -*-
# Smoke отчёта А_СравнитьОстаткиНалоговЕРПсBASБухгалтерия:
#   1) загрузка собранного .erf через COM (валидность сборки + компиляция ObjectModule);
#   2) интроспекция СКД (набор Налоги_Сравнение, поля);
#   3) COM-зеркало двух запросов отчёта на 31.12.2025 → сверка с эталонными числами промта.
import win32com.client, sys, os, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

ERF = (r"C:\Configuration_downloads\BASERP25\.claude\worktrees\sad-proskuriakova-f41f96"
       r"\_Rarzrabotki\Отчеты\А_СравнитьОстаткиНалоговЕРПсBASБухгалтерия.erf")

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')

provalov = []

# ===== 1. Загрузка .erf =====
print("=== 1. Загрузка .erf ===")
print("Файл существует:", os.path.exists(ERF), os.path.getsize(ERF) if os.path.exists(ERF) else "-", "байт")
try:
    отчет = erp.ВнешниеОтчеты.Создать(ERF, False)
    print("ВнешниеОтчеты.Создать: OK, объект:", отчет is not None)
except Exception as e:
    info = getattr(e, 'excepinfo', None)
    print("FAIL Создать:", info[2] if info else e)
    provalov.append("load")
    отчет = None

# ===== 2. Интроспекция СКД =====
print("\n=== 2. СКД ===")
try:
    схема = отчет.СхемаКомпоновкиДанных
    наборы = [схема.НаборыДанных.Получить(i).Имя for i in range(схема.НаборыДанных.Количество())]
    print("Наборы данных:", наборы)
    if наборы:
        nd = схема.НаборыДанных.Получить(0)
        поля = [nd.Поля.Получить(i).Поле for i in range(nd.Поля.Количество())]
        print("Поля набора:", поля)
        ожид = {"Организация","ВидНалога","НачальныйОстаток_ЕРП","КонечныйОстаток_ЕРП",
                "КонечныйОстаток_BuhBud","РазницаКонечныйОстаток"}
        нет = ожид - set(поля)
        print("Ключевые поля присутствуют:", "OK" if not нет else f"НЕТ: {нет}")
        if нет: provalov.append("skd_fields")
        if "Налоги_Сравнение" not in наборы: provalov.append("skd_dataset_name")
except Exception as e:
    info = getattr(e, 'excepinfo', None)
    print("WARN интроспекция СКД (не критично):", info[2] if info else e)

# ===== 3. COM-зеркало: два запроса отчёта + сверка эталона =====
print("\n=== 3. COM-зеркало (31.12.2025) ===")
НП = datetime.datetime(2025, 12, 1, 0, 0, 0)
КП = datetime.datetime(2025, 12, 31, 23, 59, 59)

# подвиды НДФЛ
qen = erp.NewObject("Запрос")
qen.Text = "ВЫБРАТЬ П.Ссылка КАК Ссылка ИЗ Перечисление.ТипыНалогов КАК П"
масНДФЛ = erp.NewObject("Массив")
for s in qen.Execute().Выгрузить():
    n = erp.XMLСтрока(s.Ссылка)
    if n.upper().startswith("НДФЛ") or n.upper().startswith("НФДЛ"):
        масНДФЛ.Добавить(s.Ссылка)

# ЕРП ПАП «Налоги»
qp = erp.NewObject("Запрос")
qp.SetParameter("НП", НП); qp.SetParameter("КП", КП); qp.SetParameter("СписокНДФЛ", масНДФЛ)
qp.Text = """
ВЫБРАТЬ Ост.Организация.КодПоЕДРПОУ КАК ЕДРПОУ,
    ВЫБОР КОГДА ВЫРАЗИТЬ(Ост.Аналитика КАК Перечисление.ТипыНалогов) В (&СписокНДФЛ)
         ТОГДА ЗНАЧЕНИЕ(Перечисление.ТипыНалогов.НДФЛ)
         ИНАЧЕ ВЫРАЗИТЬ(Ост.Аналитика КАК Перечисление.ТипыНалогов) КОНЕЦ КАК ВидНалога,
    СУММА(Ост.СуммаКонечныйОстаток) КАК КонОст
ИЗ РегистрНакопления.ПрочиеАктивыПассивы.ОстаткиИОбороты(&НП, &КП, , ,
        Статья = ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.Налоги)) КАК Ост
ГДЕ Ост.Организация.А_ВБалансе И Ост.Организация.КодПоЕДРПОУ <> ""
СГРУППИРОВАТЬ ПО Ост.Организация.КодПоЕДРПОУ,
    ВЫБОР КОГДА ВЫРАЗИТЬ(Ост.Аналитика КАК Перечисление.ТипыНалогов) В (&СписокНДФЛ)
         ТОГДА ЗНАЧЕНИЕ(Перечисление.ТипыНалогов.НДФЛ)
         ИНАЧЕ ВЫРАЗИТЬ(Ост.Аналитика КАК Перечисление.ТипыНалогов) КОНЕЦ
"""
erp_kon = {}
orgset = set()
for s in qp.Execute().Выгрузить():
    erp_kon[(s.ЕДРПОУ, erp.XMLСтрока(s.ВидНалога))] = float(s.КонОст)
    orgset.add(s.ЕДРПОУ)

эталон = {"НДС": 6668531.29, "НалогНаПрибыль": 1641963.16, "НачисленныйЕСВ": 643405.44,
          "НДФЛ": 540199.99, "ВоенныйСбор": 152888.54, "ДругиеНалоги": -232453.95}
print("ЕРП ПАП 40645273 vs эталон:")
ok = True
for vn, exp in эталон.items():
    got = erp_kon.get(("40645273", vn), 0.0)
    m = abs(got - exp) < 0.01; ok = ok and m
    print(f"  {vn:16} эталон={exp:>15,.2f} факт={got:>15,.2f} {'OK' if m else 'DIFF'}")
if not ok: provalov.append("erp_etalon")

# BuhBud 6412(НДС) для 40645273
масЕ = buh.NewObject("Массив"); масЕ.Добавить("40645273")
масС = buh.NewObject("Массив")
for c in ["6411","6412","6413","6414","6415","6417","642","651"]: масС.Добавить(c)
qb = buh.NewObject("Запрос")
qb.SetParameter("НП", НП); qb.SetParameter("КП", КП); qb.SetParameter("Список", масЕ); qb.SetParameter("Счета", масС)
qb.Text = """
ВЫБРАТЬ Ост.Счет.Код КАК Код,
   СУММА(Ост.СуммаКонечныйОстатокКт - Ост.СуммаКонечныйОстатокДт) КАК КонОст
ИЗ РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты(&НП, &КП, , , Счет.Код В (&Счета), , ) КАК Ост
ГДЕ Ост.Организация.КодПоЕДРПОУ В (&Список)
СГРУППИРОВАТЬ ПО Ост.Счет.Код
"""
buh_nds = 0.0
for s in qb.Execute().Выгрузить():
    if str(s.Код).strip() == "6412":
        buh_nds += float(s.КонОст)
m = abs(buh_nds - 5244774.70) < 0.01
print(f"BuhBud 40645273 6412(НДС) эталон=   5,244,774.70 факт={buh_nds:>15,.2f} {'OK' if m else 'DIFF'}")
if not m: provalov.append("buh_etalon")

orgs = sorted(set(k[0] for k in erp_kon))
print("\nВсе 3 орг присутствуют:", orgs, "OK" if set(orgs)>={"40645273","41597184","44590697"} else "НЕТ")
if not (set(orgs) >= {"40645273","41597184","44590697"}): provalov.append("orgs")

print("\n=== ИТОГ SMOKE:", "PASS" if not provalov else f"FAIL {provalov}", "===")
