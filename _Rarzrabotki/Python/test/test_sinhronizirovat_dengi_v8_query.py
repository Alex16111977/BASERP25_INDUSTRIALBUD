# -*- coding: utf-8 -*-
"""
v8 pre-BSL тест (Rule #-1): запрос Фазы 2 с новой колонкой ПарноеПоступление
+ проверка варианта B перевода (подчинённое Списание, Бух-документ = Надходження с UUID Поступления).

Эталон: депозит UA693395002610201537072000001, март 2026,
Списание 000Ц-000187 -> ПарноеПоступление a7f34d57-... -> BuhBud ПоступлениеНаРасчетныйСчет 00DL-9066,
поток по депозитному счёту = -44 602 752,10.
"""
import sys
from datetime import datetime

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

IBAN_DEPOZIT = "UA693395002610201537072000001"

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
buh = v8.Connect('Srvr="SQLSERVER";Ref="BuhBud";Usr="cfo";Pwd="2442"')

FAILS = 0

# --- счёт ERP по IBAN ---
q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ ПЕРВЫЕ 1 БСО.Ссылка КАК Ссылка
ИЗ Справочник.БанковскиеСчетаОрганизаций КАК БСО
ГДЕ БСО.НомерСчета = &НомерСчета
"""
q.SetParameter("НомерСчета", IBAN_DEPOZIT)
r = q.Execute().Выгрузить()
assert r.Количество() == 1, "депозитный счёт не найден в ERP"
schet = r.Получить(0).Ссылка

# --- ЗАПРОС ФАЗЫ 2 (1:1 как будет в BSL, с новой колонкой ПарноеПоступление) ---
q2 = erp.NewObject("Запрос")
q2.Text = """
ВЫБРАТЬ
	Об.Регистратор КАК Регистратор,
	ПРЕДСТАВЛЕНИЕ(Об.Регистратор) КАК РегистраторПредст,
	Об.СуммаПриход КАК Приход,
	Об.СуммаРасход КАК Расход,
	ВЫРАЗИТЬ(Об.Регистратор КАК Документ.ПоступлениеБезналичныхДенежныхСредств).ХозяйственнаяОперация КАК ХозОперацияПоступления,
	ВЫРАЗИТЬ(Об.Регистратор КАК Документ.ПоступлениеБезналичныхДенежныхСредств).А_СписаниеБезналичныхДенежныхСредств КАК ПарноеСписание,
	ВЫРАЗИТЬ(Об.Регистратор КАК Документ.СписаниеБезналичныхДенежныхСредств).ХозяйственнаяОперация КАК ХозОперацияСписания,
	ВЫРАЗИТЬ(Об.Регистратор КАК Документ.СписаниеБезналичныхДенежныхСредств).А_ПоступлениеБезналичныхДенежныхСредств КАК ПарноеПоступление,
	Соотв.УникальныйИдентификаторПриемника КАК UIDвBuhBud,
	Соотв.ТипПриемника КАК ТипПриемника,
	Соотв.ТипИсточника КАК ТипИсточника
ИЗ
	РегистрНакопления.ДенежныеСредстваБезналичные.Обороты(
		&НачалоПериода, &КонецПериода, Авто,
		БанковскийСчет = &БанковскийСчет) КАК Об
	ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.СоответствияОбъектовИнформационныхБаз КАК Соотв
	ПО Об.Регистратор = Соотв.УникальныйИдентификаторИсточника
		И (Соотв.УзелИнформационнойБазы ССЫЛКА ПланОбмена.ОбменУправлениеПредприятиемБухгалтерия20)
"""
q2.SetParameter("НачалоПериода", datetime(2026, 3, 1, 0, 0, 0))
q2.SetParameter("КонецПериода", datetime(2026, 3, 31, 23, 59, 59))
q2.SetParameter("БанковскийСчет", schet)
try:
    tz = q2.Execute().Выгрузить()
    print(f"Запрос Фазы 2 OK, рядків={tz.Количество()}")
except Exception as e:
    if hasattr(e, 'excepinfo') and e.excepinfo:
        print(f"FAIL запрос: {e.excepinfo[2]}")
    else:
        print(f"FAIL запрос: {e}")
    sys.exit(1)

# --- банк.счёт BuhBud по IBAN (как НайтиБанкСчетBuhBud: сначала UUID, потом IBAN) ---
bank_buh = None
try:
    uid_schet = erp.String(schet.УникальныйИдентификатор())
    ssylka_uid = buh.Справочники.БанковскиеСчета.ПолучитьСсылку(
        buh.NewObject("УникальныйИдентификатор", uid_schet))
    if ssylka_uid.ПолучитьОбъект() is not None:
        bank_buh = ssylka_uid
except Exception:
    pass
if bank_buh is None:
    qb = buh.NewObject("Запрос")
    qb.Text = """
    ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Справочник.БанковскиеСчета
    ГДЕ НомерСчета = &НомерСчета И Наименование <> "" И Наименование НЕ ПОДОБНО "%$$$%"
    УПОРЯДОЧИТЬ ПО Наименование
    """
    qb.SetParameter("НомерСчета", IBAN_DEPOZIT)
    rb = qb.Execute().Выгрузить()
    assert rb.Количество() >= 1, "счёт BuhBud по IBAN не найден"
    bank_buh = rb.Получить(0).Ссылка
print(f"Счёт BuhBud: {buh.String(bank_buh)}")

XML_PERECHISLENIE = "ПеречислениеДенежныхСредствНаДругойСчет"
FALLBACK_TYPES = ["СписаниеСРасчетногоСчета", "ПоступлениеНаРасчетныйСчет"]


def potok_po_schetu(ssylka_doc):
    """Поток документа BuhBud по депозитному счёту: Дт(X) - Кт(X), как ПотокДокументаПоСчетуBuhBud."""
    qp = buh.NewObject("Запрос")
    qp.Text = """
    ВЫБРАТЬ
    	СУММА(Т.Приход) - СУММА(Т.Расход) КАК Сумма
    ИЗ
    	(ВЫБРАТЬ Д.Сумма КАК Приход, 0 КАК Расход
    	ИЗ РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто(, , Регистратор = &Регистратор, , ) КАК Д
    	ГДЕ Д.СчетДт В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках))
    		И Д.СубконтоДт1 = &БанкРахунок

    	ОБЪЕДИНИТЬ ВСЕ

    	ВЫБРАТЬ 0, Д.Сумма
    	ИЗ РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто(, , Регистратор = &Регистратор, , ) КАК Д
    	ГДЕ Д.СчетКт В ИЕРАРХИИ (ЗНАЧЕНИЕ(ПланСчетов.Хозрасчетный.СчетаВБанках))
    		И Д.СубконтоКт1 = &БанкРахунок) КАК Т
    """
    qp.SetParameter("Регистратор", ssylka_doc)
    qp.SetParameter("БанкРахунок", bank_buh)
    rp = qp.Execute().Выбрать()
    if rp.Следующий() and buh.ЗначениеЗаполнено(rp.Сумма):
        return float(rp.Сумма)
    return 0.0


n_spis_perevod = 0
n_par_post = 0
n_sinhr = 0
nashel_187 = False

for i in range(tz.Количество()):
    s = tz.Получить(i)
    hoz_spis = ""
    try:
        if erp.ЗначениеЗаполнено(s.ХозОперацияСписания):
            hoz_spis = str(erp.XMLСтрока(s.ХозОперацияСписания))
    except Exception:
        pass
    if hoz_spis != XML_PERECHISLENIE:
        continue
    n_spis_perevod += 1
    predst = str(erp.String(s.РегистраторПредст))
    summa_erp = float(s.Приход) - float(s.Расход)

    if not erp.ЗначениеЗаполнено(s.ПарноеПоступление):
        print(f"  [{predst}] ПарноеПоступление ПУСТО (вариант A или несвязанный)")
        continue
    n_par_post += 1
    uid_post = str(erp.String(s.ПарноеПоступление.УникальныйИдентификатор()))

    # фоллбек-перебор типов BuhBud по UUID Поступления (как будет в BSL)
    nayden_tip = None
    doc_buh = None
    for tip in FALLBACK_TYPES:
        try:
            ssylka = getattr(buh.Документы, tip).ПолучитьСсылку(buh.NewObject("УникальныйИдентификатор", uid_post))
            if ssylka.ПолучитьОбъект() is not None:
                nayden_tip = tip
                doc_buh = ssylka
                break
        except Exception as e:
            print(f"    [{tip}] EXC: {e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else e}")
            continue
    if doc_buh is None:
        print(f"  [{predst}] FAIL: Бух-документ по UUID Поступления {uid_post} не найден ни одним типом")
        FAILS += 1
        continue

    potok = potok_po_schetu(doc_buh)
    sinhr = abs(summa_erp - potok) < 0.005
    if sinhr:
        n_sinhr += 1
    print(f"  [{predst}] ПарПост={uid_post[:8]}.. тип={nayden_tip} "
          f"Бух='{buh.String(doc_buh)}' СумЕРП={summa_erp:.2f} Поток={potok:.2f} "
          f"{'SINHR' if sinhr else 'РОЗБІЖНІСТЬ'}")

    if "000187" in predst:
        nashel_187 = True
        if not sinhr:
            print("  FAIL: 000Ц-000187 не синхронно")
            FAILS += 1
        if abs(potok - (-44602752.10)) > 0.005:
            print(f"  FAIL: поток 000187 = {potok}, ожидал -44602752.10")
            FAILS += 1
        if nayden_tip != "ПоступлениеНаРасчетныйСчет":
            print(f"  FAIL: тип 000187 = {nayden_tip}, ожидал ПоступлениеНаРасчетныйСчет")
            FAILS += 1

print(f"\nСписаний-переводов: {n_spis_perevod}, с ПарноеПоступление: {n_par_post}, синхронных: {n_sinhr}")
if not nashel_187:
    print("FAIL: 000Ц-000187 не найден в выборке")
    FAILS += 1

print("РЕЗУЛЬТАТ: " + ("OK" if FAILS == 0 else f"FAIL ({FAILS})"))
sys.exit(1 if FAILS else 0)
