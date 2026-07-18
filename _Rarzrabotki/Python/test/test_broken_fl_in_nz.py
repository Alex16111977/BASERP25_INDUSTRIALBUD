# -*- coding: utf-8 -*-
# Перевіряємо чи серед ФізОсіб у ТЧ НЗ 000Ц-000003 є битые ссылки.
# Гіпотеза: саме битая ссылка ФізЛица робить Сотрудник lookup = пустым на рядку 66.
import win32com.client, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String

# Спочатку перезаписати НЗ повністю (4 ТЧ)
qd = erp.NewObject("Запрос")
qd.Text = (
    'ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка '
    'ИЗ Документ.А_ОтражениеЗПпоКазне КАК Д '
    'ГДЕ Д.Номер = "000000001" И НЕ Д.ПометкаУдаления'
)
seld = qd.Execute().Выбрать()
seld.Следующий()
dok = seld.Ссылка.ПолучитьОбъект()
try:
    dok.СоздатьДокументНачисления()
    print("Перезаписано НЗ")
except:
    pass

# Знайти наш НЗ
q0 = erp.NewObject("Запрос")
q0.Text = (
    'ВЫБРАТЬ ПЕРВЫЕ 1 НЗ.Ссылка '
    'ИЗ Документ.НачислениеЗарплаты КАК НЗ '
    'ГДЕ НЗ.А_ДокОтражениеЗПпоКазне.Номер = "000000001" И НЕ НЗ.ПометкаУдаления '
    'УПОРЯДОЧИТЬ ПО НЗ.Дата УБЫВ'
)
sel = q0.Execute().Выбрать()
if not sel.Следующий():
    print("НЗ не знайдено")
    sys.exit(0)
нз_ref = sel.Ссылка
print(f"НЗ: {S(нз_ref)}")

obj = нз_ref.ПолучитьОбъект()

# Зібрати всі ФізОсоби з ТЧ
all_fl = {}  # uid → (ref, locations)
def add_fl(ref, loc):
    if not ref:
        return
    try:
        uid = S(ref.УникальныйИдентификатор())
    except:
        return
    if uid not in all_fl:
        all_fl[uid] = [ref, []]
    all_fl[uid][1].append(loc)

# Начисления → Сотрудник.ФизическоеЛицо
for i in range(obj.Начисления.Количество()):
    r = obj.Начисления.Получить(i)
    try:
        fl = r.Сотрудник.ФизическоеЛицо
        add_fl(fl, f"Начисления[{i+1}]")
    except:
        pass

# Удержания, НДФЛ, ВзносыФОТ — напряму ФизическоеЛицо
for тч_ім in ["Удержания", "НДФЛ", "ВзносыФОТ"]:
    тч = getattr(obj, тч_ім)
    for i in range(тч.Количество()):
        r = тч.Получить(i)
        add_fl(r.ФизическоеЛицо, f"{тч_ім}[{i+1}]")

print(f"Унікальних ФізОсіб: {len(all_fl)}")

# Перевірка битости через запит ВЫБРАТЬ ИСТИНА ИЗ Справочник.ФизическиеЛица
broken = []
for uid, (ref, locations) in all_fl.items():
    q = erp.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 1 КАК Е ИЗ Справочник.ФизическиеЛица ГДЕ Ссылка = &С"
    q.SetParameter("С", ref)
    if q.Execute().Пустой():
        broken.append((ref, locations))

print(f"\n=== БИТЫХ ССЫЛОК ФизическиеЛица: {len(broken)} ===")
for ref, locations in broken[:20]:
    print(f"  uid={S(ref.УникальныйИдентификатор())}  locations: {locations[:3]}")

# Також перевірити Сотрудников у Начисления
print("\n=== Перевірка Сотрудников у Начисления ===")
sotr_bad = 0
for i in range(obj.Начисления.Количество()):
    r = obj.Начисления.Получить(i)
    сотр = r.Сотрудник
    q = erp.NewObject("Запрос")
    q.Text = "ВЫБРАТЬ ПЕРВЫЕ 1 1 КАК Е ИЗ Справочник.Сотрудники ГДЕ Ссылка = &С"
    q.SetParameter("С", сотр)
    if q.Execute().Пустой():
        sotr_bad += 1
        print(f"  БИТЫЙ Сотрудник на рядку {i+1}: {S(сотр)}")
print(f"Всього битих Сотрудників у Начисления: {sotr_bad}")
