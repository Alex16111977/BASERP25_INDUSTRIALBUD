"""Verify acceptance после правки блока 1 ОТ в РКО.ManagerModule (Подход C).

Запуск через `mcp__python-runner__run_command`:
    python _Rarzrabotki/Python/test/test_rko_pap_ot_napravlenie_verify.py

4 проверки:
  #1 ПАП(ОТ) = 12 строк с Подр+Направ, Σ=348800
  #2 Плуги 23:59:59 ИСЧЕЗНУТ; Σ Регистратора signed=0
  #3 Σ-инвариант А_ВзСС + ПАП(ОТ) per (Орг,Подр)=0
  #4 Идемпотентность: 2-й прогон даёт идентичный результат
"""
import sys, win32com.client
from collections import defaultdict
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

ETALON_PODR = {
    "Автокран XCMG": 60800, "Виброкаток": 4900, "МАЗ 264": 46000, "МАЗ 265": 13900,
    "РДК 250": 24800, "Самосвал Камаз": 24700, "Самосвал МАЗ": 9000,
    "Телескоп 1": 28900, "Телескоп 2": 14700, "Телескоп JCB 535-140 (№4)": 70600,
    "Экскаватор  JCB 220 ": 33600, "Экскаватор  JCB 3CX 2008": 16900,
}

def find_rko(номер):
    q = erp.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ Д.Ссылка КАК Сс ИЗ Документ.РасходныйКассовыйОрдер КАК Д ГДЕ Д.Номер = &Н"
    q.SetParameter("Н", номер)
    r = q.Execute().Выгрузить()
    assert r.Количество() == 1, f"РКО {номер} не найден"
    return r[0].Сс

def repost(ref):
    # Записать с режимом Проведение — платформа сама очистит старые движения
    # регистраторных регистров (Замещение = Истина по умолчанию).
    obj = ref.ПолучитьОбъект()
    obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    print(f"  Перепроведён: {S(ref)}")

def pap_records(ref):
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ Р.Период, Р.ВидДвижения, Р.Подразделение, Р.НаправлениеДеятельности,
            Р.Статья, Р.Сумма, Р.ВидИсточника, Р.Источник
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Р
    ГДЕ Р.Регистратор = &Сс
    УПОРЯДОЧИТЬ ПО Р.Период, Р.Сумма УБЫВ
    """
    q.SetParameter("Сс", ref)
    return q.Execute().Выгрузить()

def vzss_records(ref):
    q = erp.NewObject("Запрос")
    q.Text = """
    ВЫБРАТЬ Р.Организация, Р.Подразделение, Р.ВидДвижения, Р.СуммаВзаиморасчетов КАК Сумма
    ИЗ РегистрНакопления.А_ВзаиморасчетыССотрудниками КАК Р
    ГДЕ Р.Регистратор = &Сс
    """
    q.SetParameter("Сс", ref)
    return q.Execute().Выгрузить()

ref = find_rko("N0000052986")
print("=" * 80)
print(f"Verify для РКО N0000052986")
print("=" * 80)

print("\n[STAGE 1] Первый перепровод")
repost(ref)

print("\n[STAGE 2] Acceptance проверки")

# === #1 ПАП(ОТ) = 12 строк ===
pap = pap_records(ref)
ot_rows = [r for r in pap if S(r.Статья) == "Оплата труда"]
total_ot = sum(float(r.Сумма) for r in ot_rows)
podr_set = {S(r.Подразделение): float(r.Сумма) for r in ot_rows}
napr_set = {S(r.НаправлениеДеятельности) for r in ot_rows}
assert len(ot_rows) == 12, f"#1 FAIL: ожидаем 12 строк ОТ, получили {len(ot_rows)}"
assert abs(total_ot - 348800.0) < 0.01, f"#1 FAIL: Σ ОТ != 348800: {total_ot}"
assert podr_set == ETALON_PODR, f"#1 FAIL: Подр не совпадают.\nDIFF: {set(podr_set) ^ set(ETALON_PODR)}\nGOT:  {podr_set}\nWANT: {ETALON_PODR}"
assert napr_set == {"Спецтехника"}, f"#1 FAIL: направления != Спецтехника: {napr_set}"
print(f"  ✅ #1 PASS: 12 строк ОТ, Σ=348800, Подр=эталон, Направ=Спецтехника")

# === #2 Плуги 23:59:59 ИСЧЕЗНУТ ===
plugs = [r for r in pap if S(r.Статья) in ("Вывод собственных средств", "Вложения собственных средств")]
assert len(plugs) == 0, f"#2 FAIL: найдено {len(plugs)} плуг(ов): {[(S(p.Статья), float(p.Сумма)) for p in plugs]}"
signed = sum((1 if erp.XMLСтрока(r.ВидДвижения) == "Receipt" else -1) * float(r.Сумма) for r in pap)
assert abs(signed) < 0.01, f"#2 FAIL: Σ signed по регистратору != 0: {signed}"
print(f"  ✅ #2 PASS: плугов нет, Σ signed = 0")

# === #3 Σ-инвариант А_ВзСС + ПАП(ОТ) per (Орг,Подр) ===
vzss = vzss_records(ref)
pap_by_op = defaultdict(float)
vzss_by_op = defaultdict(float)
for r in ot_rows:
    key = (S(r.Подразделение),)
    sign = 1 if erp.XMLСтрока(r.ВидДвижения) == "Receipt" else -1
    pap_by_op[key] += sign * float(r.Сумма)
for r in vzss:
    key = (S(r.Подразделение),)
    sign = 1 if erp.XMLСтрока(r.ВидДвижения) == "Receipt" else -1
    vzss_by_op[key] += sign * float(r.Сумма)
all_keys = set(pap_by_op) | set(vzss_by_op)
mismatch = []
for k in sorted(all_keys):
    d = vzss_by_op.get(k, 0) + pap_by_op.get(k, 0)
    if abs(d) > 0.01:
        mismatch.append((k, d, vzss_by_op.get(k, 0), pap_by_op.get(k, 0)))
assert not mismatch, f"#3 FAIL: Σ-инвариант не сошёлся:\n" + "\n".join(f"  {k} → Δ={d:.2f} (vzss={v}, pap={p})" for k,d,v,p in mismatch)
print(f"  ✅ #3 PASS: Σ-инвариант per (Подр) = 0 для всех {len(all_keys)} ключей")

# === #4 Идемпотентность ===
print("\n[STAGE 3] Второй перепровод (идемпотентность)")
pap_before = [(S(r.Подразделение), S(r.Статья), float(r.Сумма), erp.XMLСтрока(r.ВидДвижения)) for r in pap]
pap_before_sorted = sorted(pap_before)
repost(ref)
pap2 = pap_records(ref)
pap_after = [(S(r.Подразделение), S(r.Статья), float(r.Сумма), erp.XMLСтрока(r.ВидДвижения)) for r in pap2]
pap_after_sorted = sorted(pap_after)
assert pap_before_sorted == pap_after_sorted, f"#4 FAIL: 2-й прогон отличается:\nBEFORE: {pap_before_sorted}\nAFTER:  {pap_after_sorted}"
print(f"  ✅ #4 PASS: идемпотентность {len(pap_before)} записей")

print("\n" + "=" * 80)
print("✅ Verify PASS — все 4 acceptance выполнены")
