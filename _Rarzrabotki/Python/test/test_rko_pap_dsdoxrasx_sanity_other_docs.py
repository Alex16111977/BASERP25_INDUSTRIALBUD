"""Sanity: не-РКО документы (ОЗФУ, СБДС) не должны быть зачеплены — гард по типу документа."""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String


def find_proved_doc(тип):
    q = erp.NewObject("Запрос")
    q.Text = f"""
    ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Сс, Д.Номер
    ИЗ Документ.{тип} КАК Д
    ГДЕ Д.Проведен И НЕ Д.ПометкаУдаления
    УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ
    """
    r = q.Execute().Выгрузить()
    if r.Количество() == 0:
        return None, None
    return S(r[0].Номер), r[0].Сс


doc_types = ["ОтражениеЗарплатыВФинансовомУчете", "СписаниеБезналичныхДенежныхСредств"]


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
for тип in doc_types:
    номер, ref = find_proved_doc(тип)
    if ref is None:
        print(f"  ⚠ {тип}: проведённых не найдено")
        continue
    pap_before = pap_snap(ref)
    dsdr_before = dsdr_snap(ref)
    obj = ref.ПолучитьОбъект()
    obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    pap_after = pap_snap(ref)
    dsdr_after = dsdr_snap(ref)
    if pap_before == pap_after and dsdr_before == dsdr_after:
        print(f"  ✅ {тип} {номер}: ПАП={len(pap_before)} ДО==ПОСЛЕ, ДСДохРасх={len(dsdr_before)} ДО==ПОСЛЕ")
    else:
        failures.append((тип, номер, pap_before, pap_after, dsdr_before, dsdr_after))
        print(f"  ❌ {тип} {номер}: изменения! ПАП Δ={set(pap_after)-set(pap_before)}, ДСДохРасх Δ={set(dsdr_after)-set(dsdr_before)}")

assert not failures, f"Sanity FAIL: {len(failures)} не-РКО документов зачеплены"
print("\n✅ Sanity не-РКО PASS — наша правка не зачепила другие типы документов")
