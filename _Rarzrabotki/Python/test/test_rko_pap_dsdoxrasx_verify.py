"""Verify: 6 acceptance после правки BSL для зарплатного РКО ВКассу.

Динамический поиск тестового РКО (не хардкодим Номер — РКО мог быть перенастроен).
Берём первый РКО с ХозОп зарплатной + Ведомость=ВКассу + есть в А_Расшифровке.
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
ВЫБРАТЬ ПЕРВЫЕ 1
    Д.Ссылка КАК Сс, Д.Номер
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
assert docrows.Количество() > 0, "FAIL: не найдено зарплатных РКО с ВКассу для теста"
ref = docrows[0].Сс
print(f"Тестовый РКО: {S(docrows[0].Номер)} = {S(ref)}")

# Эталон Σ per СтрПредпр из А_ВзСС (по выплате — расход - возврат = чистая выплата per Подр)
q_vzs = erp.NewObject("Запрос")
q_vzs.Text = """
ВЫБРАТЬ
    Р.Подразделение КАК Подр,
    СУММА(ВЫБОР
            КОГДА Р.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход)
                ТОГДА Р.СуммаВзаиморасчетов
            ИНАЧЕ -Р.СуммаВзаиморасчетов
        КОНЕЦ) КАК Сум
ИЗ РегистрНакопления.А_ВзаиморасчетыССотрудниками КАК Р
ГДЕ Р.Регистратор = &Сс
СГРУППИРОВАТЬ ПО Р.Подразделение
"""
q_vzs.SetParameter("Сс", ref)
ETALON_PODR = {S(row.Подр): float(row.Сум) for row in q_vzs.Execute().Выгрузить()}
print(f"Эталон из А_ВзСС: {len(ETALON_PODR)} СтрПредпр, Σ={sum(ETALON_PODR.values()):,.2f}")
assert len(ETALON_PODR) >= 1, "FAIL: А_ВзСС эталон пустой — тест бессмысленен"
assert sum(ETALON_PODR.values()) > 0, "FAIL: Σ А_ВзСС эталона = 0 — тест бессмысленен"

# Перепровести РКО
obj = ref.ПолучитьОбъект()
obj.Записать(erp.РежимЗаписиДокумента.Проведение)
print(f"РКО {S(ref)} перепроведён.")

# Загрузить ПАП движения
q_pap = erp.NewObject("Запрос")
q_pap.Text = """
ВЫБРАТЬ Р.Период, Р.ВидДвижения, Р.Подразделение, Р.НаправлениеДеятельности, Р.Статья, Р.Сумма, Р.ВидИсточника
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Р
ГДЕ Р.Регистратор = &Сс
"""
q_pap.SetParameter("Сс", ref)
pap = q_pap.Execute().Выгрузить()
print(f"\nВсего записей ПАП: {pap.Количество()}")

# Acceptance #1: ПАП(ОТ) Приход = 12 строк per СтрПредпр == А_ВзСС
ot_app = [r for r in pap if S(r.Статья) == "Оплата труда" and erp.XMLСтрока(r.ВидДвижения) == "Receipt"]
ot_by_podr = {}
for r in ot_app:
    p = S(r.Подразделение)
    ot_by_podr[p] = ot_by_podr.get(p, 0) + float(r.Сумма)
assert len(ot_by_podr) == len(ETALON_PODR), \
    f"#1 FAIL: ПАП(ОТ) Приход {len(ot_by_podr)} строк, А_ВзСС {len(ETALON_PODR)} СтрПредпр"
mismatch1 = {k: (ot_by_podr.get(k, 0), ETALON_PODR.get(k, 0)) for k in set(ot_by_podr) | set(ETALON_PODR)
             if abs(ot_by_podr.get(k, 0) - ETALON_PODR.get(k, 0)) > 0.01}
assert not mismatch1, f"#1 FAIL: ПАП(ОТ) Приход != А_ВзСС\nПАП: {ot_by_podr}\nA_ВзСС: {ETALON_PODR}\nMismatch: {mismatch1}"
print(f"✅ #1 ПАП(ОТ) Приход = {len(ot_by_podr)} строк per СтрПредпр == А_ВзСС")

# Acceptance #2: ПАП(ДенСр) Расход = 12 строк per СтрПредпр == А_ВзСС
ds_exp = [r for r in pap if S(r.Статья) == "Денежные средства (наличные)" and erp.XMLСтрока(r.ВидДвижения) == "Expense"]
ds_by_podr = {}
for r in ds_exp:
    p = S(r.Подразделение)
    ds_by_podr[p] = ds_by_podr.get(p, 0) + float(r.Сумма)
assert len(ds_by_podr) == len(ETALON_PODR), \
    f"#2 FAIL: ПАП(ДенСр) Расход {len(ds_by_podr)} строк, А_ВзСС {len(ETALON_PODR)} СтрПредпр"
mismatch2 = {k: (ds_by_podr.get(k, 0), ETALON_PODR.get(k, 0)) for k in set(ds_by_podr) | set(ETALON_PODR)
             if abs(ds_by_podr.get(k, 0) - ETALON_PODR.get(k, 0)) > 0.01}
assert not mismatch2, f"#2 FAIL: ПАП(ДенСр) Расход != А_ВзСС\nПАП: {ds_by_podr}\nA_ВзСС: {ETALON_PODR}\nMismatch: {mismatch2}"
print(f"✅ #2 ПАП(ДенСр) Расход = {len(ds_by_podr)} строк per СтрПредпр == А_ВзСС")

# Acceptance #3: Σ ПАП per Подр = 0 (Приход ОТ - Расход ДенСр = 0)
for podr in ETALON_PODR:
    delta = ot_by_podr.get(podr, 0) - ds_by_podr.get(podr, 0)
    assert abs(delta) < 0.01, f"#3 FAIL: Δ per {podr} = {delta}"
print(f"✅ #3 Σ ПАП per Подр = 0 (баланс per подразделение сходится)")

# Acceptance #4: Плуги Вывод/Вложения СобСрс ИСЧЕЗЛИ
plugs = [r for r in pap if S(r.Статья) in ("Вывод собственных средств", "Вложения собственных средств")]
assert len(plugs) == 0, f"#4 FAIL: найдено {len(plugs)} плугов балансировки: {[(S(p.Статья), float(p.Сумма)) for p in plugs]}"
print(f"✅ #4 Плуги 23:59:59 ИСЧЕЗЛИ")

# Acceptance #5: ДСДохРасх — 21 строка per ФЛ, Подр и СтатьяДДС из А_Расшифровки
q_dsdr = erp.NewObject("Запрос")
q_dsdr.Text = """
ВЫБРАТЬ Р.Подразделение КАК Подр, Р.СтатьяДвиженияДенежныхСредств КАК СтатьяДДС,
        Р.АналитикаАктивовПассивов КАК ФЛ, Р.Сумма
ИЗ РегистрНакопления.ДвиженияДенежныеСредстваДоходыРасходы КАК Р
ГДЕ Р.Регистратор = &Сс
"""
q_dsdr.SetParameter("Сс", ref)
dsdr = q_dsdr.Execute().Выгрузить()
print(f"\nДСДохРасх: {dsdr.Количество()} строк")
# Количество ФЛ в А_Расшифровке = ожидаемое число строк ДСДохРасх
q_fl = erp.NewObject("Запрос")
q_fl.Text = """ВЫБРАТЬ КОЛИЧЕСТВО(Рш.ФизическоеЛицо) КАК N
ИЗ Документ.РасходныйКассовыйОрдер КАК Д
    ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.ВедомостьНаВыплатуЗарплатыВКассу.А_РасшифровкаВыплатыЗарплатаПоФизлицам КАК Рш
    ПО Рш.Ссылка = Д.Ведомость
ГДЕ Д.Ссылка = &Сс"""
q_fl.SetParameter("Сс", ref)
EXPECTED_FL = int(q_fl.Execute().Выгрузить()[0].N)
print(f"  Ожидаем {EXPECTED_FL} строк (= кол-во ФЛ в А_Расшифровке)")
assert dsdr.Количество() == EXPECTED_FL, f"#5 FAIL: ДСДохРасх = {dsdr.Количество()} строк (ожидаем {EXPECTED_FL})"
dsdr_podr_set = {S(r.Подр) for r in dsdr}
assert dsdr_podr_set == set(ETALON_PODR.keys()), \
    f"#5 FAIL: ДСДохРасх Подр != эталон\nПАП Подр: {dsdr_podr_set}\nЭталон: {set(ETALON_PODR.keys())}"
dsdr_dds = {S(r.СтатьяДДС) for r in dsdr}
# Для этого РКО все ФЛ из направления "Спецтехника" → одна детальная статья "Зарплата механизаторов".
# Главное: НЕ общая "Оплата за выполненные работы" из шапки РКО.
assert "Оплата за выполненные работы" not in dsdr_dds, f"#5 FAIL: всё ещё общая статья ДДС: {dsdr_dds}"
print(f"✅ #5 ДСДохРасх = 21 строка с {len(dsdr_podr_set)} СтрПредпр + детальные СтатьиДДС: {dsdr_dds}")

# Acceptance #6: Идемпотентность — 2-й перепровод даёт идентичный результат
pap_snap_before = sorted([(S(r.Подразделение), S(r.Статья), float(r.Сумма), erp.XMLСтрока(r.ВидДвижения)) for r in pap])
obj2 = ref.ПолучитьОбъект()
obj2.Записать(erp.РежимЗаписиДокумента.Проведение)
pap2 = q_pap.Execute().Выгрузить()
pap_snap_after = sorted([(S(r.Подразделение), S(r.Статья), float(r.Сумма), erp.XMLСтрока(r.ВидДвижения)) for r in pap2])
assert pap_snap_before == pap_snap_after, "#6 FAIL: 2-й прогон не идемпотентен"
print(f"✅ #6 Идемпотентность ({len(pap_snap_before)} записей ПАП)")

print("\n" + "=" * 80)
print("✅ Verify PASS — все 6 acceptance выполнены")
print("=" * 80)
