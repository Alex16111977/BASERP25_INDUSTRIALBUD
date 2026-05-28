"""Sanity: РКО без нашей ВКассу — ПАП-движения НЕ должны измениться после правки.

Запуск:
    python _Rarzrabotki/Python/test/test_rko_pap_ot_other_rko_sanity.py
"""
import sys, win32com.client
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

# Эталонные РКО — НЕ зарплатные (ХозОп=ПрочаяВыдачаДенежныхСредств), активные
# Подходят для sanity: проводятся независимо от нашего гарда на ВКассу
TEST_RKOS = ["N0000054593", "N0000054572", "N0000054568"]

def find(номер):
    q = erp.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ Д.Ссылка КАК Сс, Д.Ведомость КАК Вед ИЗ Документ.РасходныйКассовыйОрдер КАК Д ГДЕ Д.Номер = &Н"
    q.SetParameter("Н", номер)
    r = q.Execute().Выгрузить()
    if r.Количество() == 0:
        return None
    return r[0]

def pap_snap(ref):
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ Р.Период, Р.ВидДвижения, Р.Подразделение, Р.НаправлениеДеятельности,
            Р.Статья, Р.Сумма
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Р
    ГДЕ Р.Регистратор = &Сс
    УПОРЯДОЧИТЬ ПО Р.Период, Р.Сумма УБЫВ
    """
    q.SetParameter("Сс", ref)
    rows = []
    for r in q.Execute().Выгрузить():
        rows.append((str(r.Период), erp.XMLСтрока(r.ВидДвижения), S(r.Подразделение),
                     S(r.НаправлениеДеятельности), S(r.Статья), float(r.Сумма)))
    return sorted(rows)

def repost(ref):
    obj = ref.ПолучитьОбъект()
    obj.Записать(erp.РежимЗаписиДокумента.Проведение)

failures = []
for номер in TEST_RKOS:
    info = find(номер)
    if info is None:
        print(f"⚠️  {номер}: не найден — пропуск")
        continue
    ref = info.Сс
    print(f"\n--- РКО {номер} ---")
    print(f"   Ведомость: {S(info.Вед)} (заполнена={erp.ЗначениеЗаполнено(info.Вед)})")

    before = pap_snap(ref)
    print(f"   ПАП ДО: {len(before)} записей")
    repost(ref)
    after = pap_snap(ref)
    print(f"   ПАП ПОСЛЕ: {len(after)} записей")

    if before == after:
        print(f"   ✅ {номер} ПАП не изменился")
    else:
        failures.append((номер, before, after))
        print(f"   ❌ {номер} ПАП ИЗМЕНИЛСЯ")
        diff_b = set(before) - set(after)
        diff_a = set(after) - set(before)
        if diff_b:
            print(f"   Записи только в BEFORE: {diff_b}")
        if diff_a:
            print(f"   Записи только в AFTER: {diff_a}")

print("\n" + "=" * 80)
if not failures:
    print(f"✅ Sanity PASS — РКО без нашей ВКассу не зачеплены ({len(TEST_RKOS)} тестов)")
else:
    print(f"❌ Sanity FAIL — {len(failures)} из {len(TEST_RKOS)} изменились")
    sys.exit(1)
