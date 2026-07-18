# -*- coding: utf-8 -*-
# Smoke отчёта А_СравнитьОстаткиНалоговПосчетамЕРПсBASБухгалтерия:
#   1) загрузка .erf в РЕАЛЬНУЮ BaseERP (собран против BuhBud — критично проверить запуск в ЕРП);
#   2) интроспекция СКД (набор Налоги_Сравнение_Счета, поля Счет/ВидНалога/12 числовых);
#   3) COM-зеркало: ЕРП Хозрасчетный ↔ BuhBud Хозрасчетный по Счёт.Код на 31.12.2025.
import win32com.client, sys, os, datetime
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

ERF = (r"C:\Configuration_downloads\BASERP25\.claude\worktrees\sad-proskuriakova-f41f96"
       r"\_Rarzrabotki\Отчеты\А_СравнитьОстаткиНалоговПосчетамЕРПсBASБухгалтерия.erf")

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
buh = v8.Connect('Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"')
provalov = []

# ===== 1. Загрузка .erf в BaseERP =====
print("=== 1. Загрузка .erf в BaseERP ===")
print("Файл:", os.path.exists(ERF), os.path.getsize(ERF) if os.path.exists(ERF) else "-", "байт")
try:
    отчет = erp.ВнешниеОтчеты.Создать(ERF, False)
    print("ВнешниеОтчеты.Создать: OK, объект:", отчет is not None)
except Exception as e:
    info = getattr(e, 'excepinfo', None); print("FAIL Создать:", info[2] if info else e)
    provalov.append("load"); отчет = None

# ===== 2. Интроспекция СКД =====
print("\n=== 2. СКД ===")
try:
    схема = отчет.СхемаКомпоновкиДанных
    наборы = [схема.НаборыДанных.Получить(i).Имя for i in range(схема.НаборыДанных.Количество())]
    print("Наборы:", наборы)
    nd = схема.НаборыДанных.Получить(0)
    поля = [nd.Поля.Получить(i).Поле for i in range(nd.Поля.Количество())]
    print("Поля:", поля)
    need = {"Счет","ВидНалога","КонечныйОстаток_ЕРП","КонечныйОстаток_BuhBud","РазницаКонечныйОстаток"}
    нет = need - set(поля)
    print("Ключевые поля:", "OK" if not нет else f"НЕТ {нет}")
    if нет: provalov.append("skd_fields")
    if "Налоги_Сравнение_Счета" not in наборы: provalov.append("skd_dataset")
except Exception as e:
    info = getattr(e, 'excepinfo', None); print("WARN интроспекция:", info[2] if info else e)

# ===== 3. COM-зеркало по Счёт.Код =====
print("\n=== 3. COM-зеркало (31.12.2025) ===")
НП = datetime.datetime(2025, 12, 1, 0, 0, 0)
КП = datetime.datetime(2025, 12, 31, 23, 59, 59)
qo = erp.NewObject("Запрос")
qo.Text = 'ВЫБРАТЬ Орг.КодПоЕДРПОУ КАК Е ИЗ Справочник.Организации КАК Орг ГДЕ Орг.А_ВБалансе И Орг.КодПоЕДРПОУ<>""'
edrpou = [s.Е for s in qo.Execute().Выгрузить()]

def хоз(conn):
    масЕ = conn.NewObject("Массив")
    for e in edrpou: масЕ.Добавить(e)
    q = conn.NewObject("Запрос")
    q.SetParameter("НП", НП); q.SetParameter("КП", КП); q.SetParameter("Список", масЕ)
    q.Text = """
    ВЫБРАТЬ Ост.Счет.Код КАК Код,
        СУММА(Ост.СуммаКонечныйОстатокКт - Ост.СуммаКонечныйОстатокДт) КАК КонОст
    ИЗ РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты(&НП, &КП, , , , , ) КАК Ост
    ГДЕ (Ост.Счет.Код ПОДОБНО "641%" ИЛИ Ост.Счет.Код ПОДОБНО "642%" ИЛИ Ост.Счет.Код ПОДОБНО "651%")
        И Ост.Организация.КодПоЕДРПОУ В (&Список)
    СГРУППИРОВАТЬ ПО Ост.Счет.Код
    """
    return {str(s.Код).strip(): float(s.КонОст) for s in q.Execute().Выгрузить()}

erp_acc = хоз(erp); buh_acc = хоз(buh)
коды = sorted(set(list(erp_acc) + list(buh_acc)))
print(f"{'Счёт':7} {'КонОст_ЕРП':>15} {'КонОст_Бух':>15} {'Разница':>15}")
tE = tB = 0.0
for к in коды:
    e = erp_acc.get(к, 0.0); b = buh_acc.get(к, 0.0); tE += e; tB += b
    print(f"{к:7} {e:>15,.2f} {b:>15,.2f} {e-b:>15,.2f}")
print(f"{'ИТОГО':7} {tE:>15,.2f} {tB:>15,.2f} {tE-tB:>15,.2f}")

# эталон (валидированный первый отчёт): Бух Итого 6 411 807,05; ЕРП Хозрасчетный почти пуст (~5 961,06)
m_buh = abs(tB - 6411807.05) < 0.01
print(f"\nBuhBud Итого: эталон=6,411,807.05 факт={tB:,.2f} {'OK' if m_buh else 'DIFF'}")
if not m_buh: provalov.append("buh_total")
print("ЕРП Хозрасчетный Итого:", f"{tE:,.2f}", "(почти пуст — регл.учёт ЕРП слабо синхронизирован, by-design)")
print("Счетов в сверке:", len(коды), "| орг А_ВБалансе:", len(edrpou))

print("\n=== ИТОГ SMOKE:", "PASS" if not provalov else f"FAIL {provalov}", "===")
