# -*- coding: utf-8 -*-
# Тест процедури ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП
# ПЕРЕДУМОВА: код ObjectModule.bsl завантажений у конфігурацію (Конфігуратор → Оновити конфігурацію БД)
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
conn = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = conn.String
VZ = conn.ЗначениеЗаполнено

q = conn.NewObject("Запрос")
q.Text = "ВЫБРАТЬ Док.Ссылка ИЗ Документ.ОтражениеЗарплатыВФинансовомУчете КАК Док ГДЕ Док.Номер = &Н И НЕ Док.ПометкаУдаления"
q.SetParameter("Н", "000Ц-000001")
sel = q.Execute().Choose()
sel.Next()
obj = sel.Ссылка.GetObject()

print(f"Документ: {S(sel.Ссылка)}")
print(f"Період: {S(obj.ПериодРегистрации)}")

# ДО виклику
kaz_before = obj.А_РаспределениеКазна.Count()
print(f"\nДО виклику: А_РаспределениеКазна = {kaz_before} рядків")

# Рахуємо дублікати для Петрова (було 54, має стати ~4)
petrov_before = 0
for i in range(kaz_before):
    r = obj.А_РаспределениеКазна.Get(i)
    if "Петров Михайло" in S(r.ФИО):
        petrov_before += 1
print(f"  Петров Михайло ДО: {petrov_before} рядків")

# Стецюк ДО
stetsuk_before_docs = set()
stetsuk_before_sum = 0
for i in range(kaz_before):
    r = obj.А_РаспределениеКазна.Get(i)
    if "Стецюк" in S(r.ФИО):
        stetsuk_before_docs.add(S(r.ДокРаспределениеЗП))
        stetsuk_before_sum += float(S(r.Сумма).replace('\xa0', '').replace(' ', '').replace(',', '.'))
print(f"  Стецюк ДО: {stetsuk_before_sum:.2f}, документи: {stetsuk_before_docs}")

# ВИКЛИК процедури
print("\n>>> Виклик ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()...")
try:
    obj.ЗаполнитьОтражениеЗарплатыВФинансовомУчетеИзБазЗП()
    print(">>> OK")
except Exception as e:
    print(f">>> ПОМИЛКА: {e}")
    sys.exit(1)

# ПІСЛЯ
kaz_after = obj.А_РаспределениеКазна.Count()
print(f"\nПІСЛЯ виклику: А_РаспределениеКазна = {kaz_after} рядків")

petrov_after = 0
for i in range(kaz_after):
    r = obj.А_РаспределениеКазна.Get(i)
    if "Петров Михайло" in S(r.ФИО):
        petrov_after += 1
print(f"  Петров Михайло ПІСЛЯ: {petrov_after} рядків")

stetsuk_after_docs = set()
stetsuk_after_sum = 0
stetsuk_details = []
for i in range(kaz_after):
    r = obj.А_РаспределениеКазна.Get(i)
    if "Стецюк" in S(r.ФИО):
        stetsuk_after_docs.add(S(r.ДокРаспределениеЗП))
        sum_val = float(S(r.Сумма).replace('\xa0', '').replace(' ', '').replace(',', '.'))
        stetsuk_after_sum += sum_val
        stetsuk_details.append({
            "стаття": S(r.СтатьяДвиженияДенежныхСредств)[:35],
            "док": S(r.ДокРаспределениеЗП)[:60],
            "підр": S(r.Подразделение)[:20],
            "сум": sum_val,
            "ІНН": S(r.ИНН),
            "КодКазна": S(r.КодСотрудникаКазна),
        })

print(f"\n  Стецюк ПІСЛЯ: {stetsuk_after_sum:.2f}")
print(f"  Документи:")
for d in stetsuk_after_docs:
    print(f"    - {d}")
print(f"  Рядки:")
for d in stetsuk_details:
    print(f"    [{d['стаття']:35s}] {d['док'][:50]:50s} Підр={d['підр']:20s} Сум={d['сум']:>12.2f} ІНН={d['ІНН']} КодКазна={d['КодКазна']}")

# Перевірки
print(f"\n=== Перевірки ===")
print(f"  Петров: {petrov_before} → {petrov_after} (було 54, має бути ~4)")
print(f"  Стецюк сума: {stetsuk_before_sum:.2f} → {stetsuk_after_sum:.2f}")
print(f"  Стецюк очікувана сума за грудень: 36046.63 (ЗП №268 = 16346.63 + ф2 №33 = 19700)")

# Начислення
afiz = obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Count()
anach = obj.НачисленнаяЗарплатаИВзносы.Count()
andfl = obj.НачисленныйНДФЛ.Count()
empty_fl = sum(1 for i in range(afiz) if not VZ(obj.НачисленнаяЗарплатаИВзносыПоФизлицам.Get(i).ФизическоеЛицо))
empty_dept = sum(1 for i in range(anach) if not VZ(obj.НачисленнаяЗарплатаИВзносы.Get(i).ПодразделениеПредприятия))
empty_sotr = sum(1 for i in range(kaz_after) if not VZ(obj.А_РаспределениеКазна.Get(i).Сотрудник))

print(f"\n  НачФизлицам: {afiz}, порожніх ФізЛиц: {empty_fl}")
print(f"  Начислення: {anach}, порожніх Підрозділів: {empty_dept}")
print(f"  НДФЛ: {andfl}")
print(f"  Порожніх Сотрудників в А_РаспределениеКазна: {empty_sotr}")

if abs(stetsuk_after_sum - 36046.63) < 1:
    print("\n*** Стецюк сума співпадає з очікуваною ✅ ***")
elif abs(stetsuk_after_sum - 16346.63) < 1:
    print("\n*** Стецюк = 16346.63 — старий код (без ф2). Конфігурація не завантажена в базу! ***")
else:
    print(f"\n*** Невідома сума Стецюк: {stetsuk_after_sum:.2f} ***")
