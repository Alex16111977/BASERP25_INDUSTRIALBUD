# -*- coding: utf-8 -*-
# GREEN смок: перезаповнити НачислениеЗарплаты 000Ц-000005 НОВИМ кодом (1 документ) і перевірити:
#   - больничні Начисления = GROSS, без мінусу;
#   - ЗарплатаКВыплате нетто по больничним == take-home якір (Григоренко 4135.33 / Горбачова 17543.40 / Ямковий 10721.46).
# ПЕРЕДУМОВА: оновлений ObjectModule.bsl завантажений у конфігурацію (F7 / db-load-xml + db-update).
import win32com.client, sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String


def num(x):
    return float(S(x).replace('\xa0', '').replace(' ', '').replace(',', '.') or "0")


def run(text, **p):
    q = erp.NewObject("Запрос"); q.Text = text
    for k, vv in p.items():
        q.SetParameter(k, vv)
    return q.Execute().Выгрузить()


parent_ref = run("ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Р ИЗ Документ.А_ОтражениеЗПпоКазне КАК Д ГДЕ Д.Номер=&Н", Н="000000002").Получить(0).Р
obj = parent_ref.ПолучитьОбъект()

print(">>> СоздатьДокументНачисления() — перезапис 000Ц-000005 новим кодом ...")
try:
    obj.СоздатьДокументНачисления()
    print(">>> OK\n")
except Exception as e:
    ei = getattr(e, 'excepinfo', None)
    print(f">>> ПОМИЛКА: {ei[2] if ei else e}")
    sys.exit(1)

nz = run("ВЫБРАТЬ ПЕРВЫЕ 1 НЗ.Ссылка КАК Р, НЗ.Проведен КАК П ИЗ Документ.НачислениеЗарплаты КАК НЗ "
         "ГДЕ НЗ.А_ДокОтражениеЗПпоКазне=&О И НЕ НЗ.ПометкаУдаления", О=parent_ref)
nzref = nz.Получить(0).Р
print(f"НачислениеЗарплаты Проведен={S(nz.Получить(0).П)}")

TARGETS = ["Григоренко Олександр", "Горбачова", "Ямковий"]
ANCHOR_TAKEHOME = {"Григоренко Олександр": 4135.33, "Горбачова": 17543.40, "Ямковий": 10721.46}
ANCHOR_GROSS = {"Григоренко Олександр": 5370.57, "Горбачова": 17543.40, "Ямковий": 13923.97}

print("\n=== Начисления больничних (після фіксу) ===")
t = run("ВЫБРАТЬ Т.Сотрудник.ФизическоеЛицо.Наименование КАК ФЛ, Т.Начисление КАК Вид, "
        "Т.Подразделение.Наименование КАК Подр, Т.Результат КАК Рез "
        "ИЗ Документ.НачислениеЗарплаты.Начисления КАК Т ГДЕ Т.Ссылка=&Р УПОРЯДОЧИТЬ ПО ФЛ", Р=nzref)
gross_now = {}
for i in range(t.Количество()):
    r = t.Получить(i); fl = S(r.ФЛ)
    if any(x in fl for x in TARGETS):
        key = next(x for x in TARGETS if x in fl)
        gross_now[key] = gross_now.get(key, 0) + num(r.Рез)
        print(f"  {fl[:32]:32s} | {S(r.Вид)[:20]:20s} | {S(r.Подр)[:20]:20s} | {num(r.Рез):>10.2f}")

print("\n=== ЗарплатаКВыплате нетто по больничним ===")
t = run("ВЫБРАТЬ Р.ФизическоеЛицо.Наименование КАК ФЛ, "
        "СУММА(ВЫБОР КОГДА Р.ВидДвижения=ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход) ТОГДА Р.СуммаКВыплате ИНАЧЕ -Р.СуммаКВыплате КОНЕЦ) КАК Нетто "
        "ИЗ РегистрНакопления.ЗарплатаКВыплате КАК Р ГДЕ Р.Регистратор=&Д "
        "СГРУППИРОВАТЬ ПО Р.ФизическоеЛицо.Наименование", Д=nzref)
pay_now = {}
for i in range(t.Количество()):
    r = t.Получить(i); fl = S(r.ФЛ)
    if any(x in fl for x in TARGETS):
        key = next(x for x in TARGETS if x in fl)
        pay_now[key] = pay_now.get(key, 0) + num(r.Нетто)

ok = True
print(f"\n{'ФЛ':24s} {'GROSS факт':>11s} {'GROSS якір':>11s} | {'Выплата факт':>13s} {'take-home якір':>14s}")
for key in TARGETS:
    g = gross_now.get(key, 0); ga = ANCHOR_GROSS[key]
    p = pay_now.get(key, 0); pa = ANCHOR_TAKEHOME[key]
    g_ok = abs(g - ga) <= 0.02; p_ok = abs(p - pa) <= 0.02
    ok = ok and g_ok and p_ok
    print(f"{key:24s} {g:>11.2f} {ga:>11.2f} [{'OK' if g_ok else 'FAIL'}] | {p:>13.2f} {pa:>14.2f} [{'OK' if p_ok else 'FAIL'}]")

print("\n" + ("==> SMOKE GREEN PASS" if ok else "==> SMOKE FAIL"))
sys.exit(0 if ok else 1)
