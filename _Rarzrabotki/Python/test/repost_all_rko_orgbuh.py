"""Массовое перепроведение всех проведённых РКО с заполненным А_ОрганизацияБухгалтерия,
чтобы движения А_ВзСС стали Форма1 + ОргБух. Идемпотентно (ОбработкаПроведения чистит
и переписывает движения). Проверяет результат per документ.
"""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ Д.Ссылка КАК Ссылка, Д.Номер КАК Номер,
       ПРЕДСТАВЛЕНИЕ(Д.А_ОрганизацияБухгалтерия) КАК ОргБух
ИЗ Документ.РасходныйКассовыйОрдер КАК Д
ГДЕ Д.А_ОрганизацияБухгалтерия <> ЗНАЧЕНИЕ(Справочник.Организации.ПустаяСсылка) И Д.Проведен
УПОРЯДОЧИТЬ ПО Д.Номер
"""
docs = q.Execute().Выгрузить()
print(f"РКО с ОргБух (проведённых): {docs.Количество()}")

ok = 0
errors = []
for d in docs:
    num = S(d.Номер)
    try:
        obj = d.Ссылка.ПолучитьОбъект()
        obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    except Exception as e:
        msg = e.excepinfo[2] if hasattr(e, 'excepinfo') and e.excepinfo else str(e)
        errors.append(f"{num}: repost FAIL: {msg}")
        continue
    # проверка движений
    qv = erp.NewObject("Запрос")
    qv.Text = """
    ВЫБРАТЬ ПРЕДСТАВЛЕНИЕ(Р.ФормаPL) КАК ФормаPL, ПРЕДСТАВЛЕНИЕ(Р.ОрганизацияБухгалтерия) КАК ОргБух,
           СУММА(Р.СуммаВзаиморасчетов) КАК Сумма
    ИЗ РегистрНакопления.А_ВзаиморасчетыССотрудниками КАК Р
    ГДЕ Р.Регистратор = &Док
    СГРУППИРОВАТЬ ПО ПРЕДСТАВЛЕНИЕ(Р.ФормаPL), ПРЕДСТАВЛЕНИЕ(Р.ОрганизацияБухгалтерия)
    """
    qv.SetParameter("Док", d.Ссылка)
    rows = qv.Execute().Выгрузить()
    forms = {S(r.ФормаPL) for r in rows}
    obs = {S(r.ОргБух) for r in rows}
    total = sum(float(r.Сумма) for r in rows)
    bad = []
    if rows.Количество() and forms != {"Форма1"}:
        bad.append(f"ФормаPL={forms}")
    if rows.Количество() and obs != {S(d.ОргБух)}:
        bad.append(f"ОргБух={obs}!={S(d.ОргБух)}")
    if bad:
        errors.append(f"{num}: " + "; ".join(bad))
    else:
        ok += 1
        print(f"  ✓ {num}: Форма1, ОргБух={S(d.ОргБух)}, Σ={total:,.2f}")

print("\n" + "=" * 80)
if errors:
    print(f"⚠ перепроведено OK={ok}, проблемы={len(errors)}:")
    for e in errors:
        print("   - " + e)
    sys.exit(1)
print(f"✅ DONE — {ok} РКО перепроведено, все А_ВзСС = Форма1 + соответствующий ОргБух")
print("=" * 80)
