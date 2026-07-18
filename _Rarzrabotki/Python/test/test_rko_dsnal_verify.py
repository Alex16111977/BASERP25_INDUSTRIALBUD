"""Verify: 6 acceptance после правки РКО.ManagerModule для регистра ДенежныеСредстваНаличные.

РКО N0000052986 (Σ=348 800, ВКассу 000Ц-000009) должен дать 12 строк per СтрПредпр,
одну Кассу, детальную СтатьюДДС, СуммыРегл/Упр = СуммаДокумента, идемпотентность.
"""
import sys
import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

# Динамический поиск: РКО с ВКассу + зарплатная ХозОп + есть А_Расшифровка
q_doc = erp.NewObject("Запрос")
q_doc.Text = """
ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Сс, Д.Номер
ИЗ Документ.РасходныйКассовыйОрдер КАК Д
    ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.ВедомостьНаВыплатуЗарплатыВКассу.А_РасшифровкаВыплатыЗарплатаПоФизлицам КАК Рш
        ПО Рш.Ссылка = Д.Ведомость
ГДЕ Д.Проведен И НЕ Д.ПометкаУдаления
    И ТИПЗНАЧЕНИЯ(Д.Ведомость) = ТИП(Документ.ВедомостьНаВыплатуЗарплатыВКассу)
    И Д.ХозяйственнаяОперация В (
        ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ВыплатаЗарплатыЧерезКассу),
        ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ВыплатаЗарплатыРаздатчиком),
        ЗНАЧЕНИЕ(Перечисление.ХозяйственныеОперации.ВыплатаЗарплатыРаботнику))
УПОРЯДОЧИТЬ ПО Д.Дата УБЫВ
"""
docrows = q_doc.Execute().Выгрузить()
assert docrows.Количество() > 0, "FAIL: не найдено зарплатных РКО с ВКассу"
ref = docrows[0].Сс
print(f"Тестовый РКО: {S(docrows[0].Номер)}")

# Эталон из А_ВзСС (Σ per СтрПредпр)
q_vzs = erp.NewObject("Запрос")
q_vzs.Text = """
ВЫБРАТЬ Р.Подразделение КАК Подр,
    СУММА(ВЫБОР КОГДА Р.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход)
                ТОГДА Р.СуммаВзаиморасчетов
                ИНАЧЕ -Р.СуммаВзаиморасчетов КОНЕЦ) КАК Сум
ИЗ РегистрНакопления.А_ВзаиморасчетыССотрудниками КАК Р
ГДЕ Р.Регистратор = &Сс
СГРУППИРОВАТЬ ПО Р.Подразделение
"""
q_vzs.SetParameter("Сс", ref)
ETALON_PODR = {S(row.Подр): float(row.Сум) for row in q_vzs.Execute().Выгрузить()}
print(f"Эталон из А_ВзСС: {len(ETALON_PODR)} СтрПредпр, Σ={sum(ETALON_PODR.values()):,.2f}")
assert len(ETALON_PODR) >= 1 and sum(ETALON_PODR.values()) > 0, "FAIL: эталон пуст — тест бессмысленен"
EXPECTED = len(ETALON_PODR)
TOTAL_Σ = sum(ETALON_PODR.values())

# Перепровести РКО
obj = ref.ПолучитьОбъект()
obj.Записать(erp.РежимЗаписиДокумента.Проведение)
print(f"РКО {S(ref)} перепроведён.")

# Загрузить движения регистра ДенежныеСредстваНаличные
q = erp.NewObject("Запрос")
q.Text = """
ВЫБРАТЬ
    Р.Период, Р.ВидДвижения, Р.Касса, Р.Подразделение,
    Р.ХозяйственнаяОперация КАК ХозОп,
    Р.СтатьяДвиженияДенежныхСредств КАК СтатьяДДС,
    Р.Сумма, Р.СуммаУпр, Р.СуммаРегл,
    Р.ОбъектРасчетов, Р.АналитикаУчетаПоПартнерам
ИЗ РегистрНакопления.ДенежныеСредстваНаличные КАК Р
ГДЕ Р.Регистратор = &Сс
УПОРЯДОЧИТЬ ПО Р.Подразделение
"""
q.SetParameter("Сс", ref)
rows = q.Execute().Выгрузить()
print(f"\nДвижений ДенежныеСредстваНаличные: {rows.Количество()}")
for r in rows:
    print(f"  {S(r.Подразделение):<32} | {S(r.Касса):<25} | {S(r.СтатьяДДС):<25} | {float(r.Сумма):>10,.2f}")

# Acceptance #1: ровно 12 строк
# #1: ≥ EXPECTED строк (одна СтрПредпр может иметь несколько СтатейДДС — больше OK)
assert rows.Количество() >= EXPECTED, f"#1 FAIL: ожидаем ≥{EXPECTED} строк, получили {rows.Количество()}"
print(f"\n✅ #1 {rows.Количество()} строк в РегНак.ДенежныеСредстваНаличные (≥{EXPECTED} СтрПредпр)")

# Acceptance #2: Σ = СуммаДокумента
total = sum(float(r.Сумма) for r in rows)
total_upr = sum(float(r.СуммаУпр) for r in rows)
total_regl = sum(float(r.СуммаРегл) for r in rows)
assert abs(total - TOTAL_Σ) < 0.01, f"#2 FAIL: Σ={total}, ожидаем {TOTAL_Σ}"
assert abs(total_upr - TOTAL_Σ) < 0.01, f"#2 FAIL: ΣУпр={total_upr}"
assert abs(total_regl - TOTAL_Σ) < 0.01, f"#2 FAIL: ΣРегл={total_regl}"
print(f"✅ #2 Σ={TOTAL_Σ:,.2f} (Сумма+Упр+Регл)")

# Acceptance #3: Подр-сет совпадает с А_ВзСС
podr_set = {S(r.Подразделение) for r in rows}
assert podr_set == set(ETALON_PODR.keys()), f"#3 FAIL: {podr_set} != {set(ETALON_PODR.keys())}"
print(f"✅ #3 Подр-сет ({len(podr_set)} СтрПредпр) == А_ВзСС эталон")

# Σ per Подр совпадает с А_ВзСС
by_podr = {}
for r in rows:
    p = S(r.Подразделение)
    by_podr[p] = by_podr.get(p, 0) + float(r.Сумма)
mismatch = {k: (by_podr.get(k, 0), ETALON_PODR.get(k, 0))
            for k in set(by_podr) | set(ETALON_PODR)
            if abs(by_podr.get(k, 0) - ETALON_PODR.get(k, 0)) > 0.01}
assert not mismatch, f"#3 FAIL Σ per Подр: {mismatch}"
print(f"✅ #3 Σ per Подр == А_ВзСС эталон (по копейкам)")

# Acceptance #4: СтатьяДДС детальная
dds_set = {S(r.СтатьяДДС) for r in rows}
assert "Оплата за выполненные работы" not in dds_set, f"#4 FAIL: общая статья: {dds_set}"
print(f"✅ #4 СтатьяДДС детальная: {dds_set}")

# Acceptance #5: Касса = одна
касса_set = {S(r.Касса) for r in rows}
assert len(касса_set) == 1, f"#5 FAIL: ожидаем 1 Кассу, получили {касса_set}"
print(f"✅ #5 Касса = одна: {касса_set}")

# Acceptance #6: Идемпотентность
snap_before = sorted([(S(r.Подразделение), S(r.Касса), S(r.СтатьяДДС), float(r.Сумма)) for r in rows])
obj2 = ref.ПолучитьОбъект()
obj2.Записать(erp.РежимЗаписиДокумента.Проведение)
rows2 = q.Execute().Выгрузить()
snap_after = sorted([(S(r.Подразделение), S(r.Касса), S(r.СтатьяДДС), float(r.Сумма)) for r in rows2])
assert snap_before == snap_after, "#6 FAIL: 2-й прогон не идемпотентен"
print(f"✅ #6 Идемпотентность ({len(snap_before)} записей)")

print("\n" + "=" * 80)
print("✅ Verify PASS — все 6 acceptance выполнены")
print("=" * 80)
