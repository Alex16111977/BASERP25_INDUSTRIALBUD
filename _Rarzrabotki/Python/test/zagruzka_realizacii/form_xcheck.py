import re, sys, collections
sys.stdout.reconfigure(encoding="utf-8")
K = sys.argv[1]
x = open(K + r"\Forms\Форма\Ext\Form.xml", encoding="utf-8-sig").read()
m = open(K + r"\Forms\Форма\Ext\Form\Module.bsl", encoding="utf-8-sig").read()
names = set(re.findall(r'<\w+ name="([^"]+)" id="', x))
procs = set(re.findall(r'^(?:Процедура|Функция) (\w+)\(', m, re.M))
attrs = set(re.findall(r'<Attribute name="([^"]+)"', x))
cols = set(re.findall(r'<Column name="([^"]+)"', x))
bad = []
for e in sorted(set(re.findall(r'Элементы\.(\w+)', m))):
    if e not in names: bad.append("Элементы." + e)
for h in re.findall(r'<Event name="\w+">(\w+)</Event>', x):
    if h not in procs: bad.append("обробник " + h)
for a in re.findall(r'<Action>(\w+)</Action>', x):
    if a not in procs: bad.append("дія " + a)
for dp in re.findall(r'<DataPath>([^<]+)</DataPath>', x):
    head = dp.split(".")[0]
    if head not in attrs: bad.append("DataPath " + dp)
    if head in ("Распознавание", "Варианты", "Кандидаты") and dp.count(".") == 1 and dp.split(".")[1] not in cols:
        bad.append("колонка " + dp)
# id унікальні в межах елементів
ids = re.findall(r'<(?:\w+) name="[^"]+" id="(-?\d+)"', x.split("<Attributes>")[0])
dup = [k for k, v in collections.Counter(ids).items() if v > 1]
if dup: bad.append("дубль id елементів " + ",".join(dup))
# реквізити форми, на які посилається модуль (присвоєння верхнього рівня)
for v in sorted(set(re.findall(r'^\t+(\w+) = ', m, re.M))):
    pass
# заголовки без uk
for blk in re.findall(r'<Title(?: formatted="false")?>(.*?)</Title>', x, re.S):
    if "<v8:lang>uk</v8:lang>" not in blk: bad.append("Title без uk")
# ЗначениеВРеквизитФормы вимагає однаковий набір колонок у ТЗ модуля об'єкта і в реквізиті форми
# (09.10: зайва колонка «Группа» у реквізиті -> «Недопустимое значение параметра (параметр номер '2')»)
om = open(K + r"\Ext\ObjectModule.bsl", encoding="utf-8-sig").read()
for атр, функ in (("Распознавание", "НоваяТаблицаРаспознавания"), ("Варианты", "НоваяТаблицаВариантов")):
    blk = re.search(rf'<Attribute name="{атр}" id="\d+">(.*?)</Attribute>', x, re.S)
    fb = re.search(rf'Функция {функ}\(\)(.*?)КонецФункции', om, re.S)
    if blk and fb:
        a = set(re.findall(r'<Column name="([^"]+)"', blk.group(1)))
        b = set(re.findall(r'Колонки\.Добавить\("([^"]+)"', fb.group(1)))
        if a != b:
            bad.append(f"колонки {атр} ≠ {функ}: {sorted(a ^ b)}")
print("ПРОБЛЕМИ:", bad or "немає", "| елементів", len(names), "| процедур", len(procs))
