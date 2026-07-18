"""Sanity: РКО БЕЗ нашей ВКассу не должны быть зачеплены."""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

# Найти 3 контрольных РКО — без зарплатной ХозОп
q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ ПЕРВЫЕ 3 Д.Ссылка КАК Сс, Д.Номер, Д.ХозяйственнаяОперация КАК ХозОп
ИЗ Документ.РасходныйКассовыйОрдер КАК Д
ГДЕ Д.Проведен И НЕ Д.ПометкаУдаления
    И Д.ХозяйственнаяОперация = ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ПрочаяВыдачаДенежныхСредств)
УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ
"""
controls = [(S(r.Номер), r.Сс) for r in q.Execute().Выгрузить()]
print(f"Контрольные РКО (не зарплатные): {[c[0] for c in controls]}")


def pap_snap(ref_):
    qp = erp.NewObject("Запрос")
    qp.Text = """
    ВЫБРАТЬ Р.Период, Р.ВидДвижения, Р.Подразделение, Р.НаправлениеДеятельности, Р.Статья, Р.Сумма
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Р
    ГДЕ Р.Регистратор = &Сс
    УПОРЯДОЧИТЬ ПО Р.Период, Р.Сумма УБЫВ
    """
    qp.SetParameter("Сс", ref_)
    return sorted([(str(r.Период), erp.XMLСтрока(r.ВидДвижения), S(r.Подразделение),
                    S(r.НаправлениеДеятельности), S(r.Статья), float(r.Сумма))
                   for r in qp.Execute().Выгрузить()])


def dsdr_snap(ref_):
    qd = erp.NewObject("Запрос")
    qd.Text = """
    ВЫБРАТЬ Р.Период, Р.Подразделение, Р.СтатьяДвиженияДенежныхСредств, Р.Сумма
    ИЗ РегистрНакопления.ДвиженияДенежныеСредстваДоходыРасходы КАК Р
    ГДЕ Р.Регистратор = &Сс
    УПОРЯДОЧИТЬ ПО Р.Период, Р.Сумма УБЫВ
    """
    qd.SetParameter("Сс", ref_)
    return sorted([(str(r.Период), S(r.Подразделение), S(r.СтатьяДвиженияДенежныхСредств), float(r.Сумма))
                   for r in qd.Execute().Выгрузить()])


failures = []
for номер, ref in controls:
    pap_before = pap_snap(ref)
    dsdr_before = dsdr_snap(ref)
    obj = ref.ПолучитьОбъект()
    obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    pap_after = pap_snap(ref)
    dsdr_after = dsdr_snap(ref)
    if pap_before == pap_after and dsdr_before == dsdr_after:
        print(f"  ✅ {номер}: ПАП={len(pap_before)} ДО==ПОСЛЕ, ДСДохРасх={len(dsdr_before)} ДО==ПОСЛЕ")
    else:
        failures.append((номер, pap_before, pap_after, dsdr_before, dsdr_after))
        print(f"  ❌ {номер}: изменения! ПАП Δ={set(pap_after)-set(pap_before)}, ДСДохРасх Δ={set(dsdr_after)-set(dsdr_before)}")

assert not failures, f"Sanity FAIL: {len(failures)} РКО зачеплены"
print("\n✅ Sanity РКО PASS — 3 РКО без ВКассу не зачеплены")
