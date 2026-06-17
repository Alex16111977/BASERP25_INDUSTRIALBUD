# -*- coding: utf-8 -*-
# Трансформер Form.xml обработки СинхронизироватьВзаиморасчеты: товары->взаиморасчёты.
import re, io, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

PATH = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\silly-colden-130551\_Rarzrabotki\Обработки\СинхронизироватьВзаиморасчеты\Forms\Форма\Ext\Form.xml"
with io.open(PATH, encoding="utf-8") as f:
    t = f.read()

def rep(old, new, n=None):
    global t
    c = t.count(old)
    assert c > 0, "не найдено: " + old
    if n is not None:
        assert c == n, "ожидалось %d вхождений '%s', найдено %d" % (n, old, c)
    t = t.replace(old, new)

# --- тип основного реквизита формы ---
rep("cfg:ExternalDataProcessorObject.СинхронизироватьТоварыТолькоТовары",
    "cfg:ExternalDataProcessorObject.СинхронизироватьВзаиморасчеты", 1)

# --- DataPath: ТЧ ТаблицаСкладов -> ТаблицаКонтрагентов ---
rep("Объект.ТаблицаСкладов", "Объект.ТаблицаКонтрагентов")
# имя таблицы формы и все ссылки (AdditionSource, ContextMenu, обработчики)
rep("ТаблицаСкладовФорма", "ТаблицаКонтрагентовФорма")
# подполя мастера
rep("Объект.ТаблицаКонтрагентов.Склад", "Объект.ТаблицаКонтрагентов.Контрагент")
rep("Объект.ТаблицаКонтрагентов.Позиций", "Объект.ТаблицаКонтрагентов.Договоров")  # + Total...

# --- ТЧ Документов: подполя ---
rep("Объект.ТаблицаДокументов.Номенклатура", "Объект.ТаблицаДокументов.Контрагент")
rep("Объект.ТаблицаДокументов.КоличествоЕРП", "Объект.ТаблицаДокументов.СуммаЕРП")
rep("Объект.ТаблицаДокументов.КоличествоБух", "Объект.ТаблицаДокументов.СуммаБух")

# --- объектные реквизиты-фильтры ---
rep("Объект.Номенклатура", "Объект.Контрагент")   # InputField DataPath + Save
rep("Объект.Склад", "Объект.Договор")

# --- команда АнализТекущегоСчета -> АнализТекущегоКонтрагента ---
rep("АнализТекущегоСчета", "АнализТекущегоКонтрагента")

# --- имена элементов-колонок (косметика, но чистим) ---
rep("ТССклад", "ТСКонтрагент")
rep("ТСПозицій", "ТСДоговорів")
rep("ТДНоменклатура", "ТДКонтрагент")
rep("ТДКоличествоЕРП", "ТДСуммаЕРП")
rep("ТДКоличествоБух", "ТДСуммаБух")

# --- заголовки (titles) ---
# "Номенклатура" в title всегда -> "Контрагент" (и фильтр, и колонка)
t = t.replace("<v8:content>Номенклатура</v8:content>", "<v8:content>Контрагент</v8:content>")
# "Склад" title: фильтр (id=33) -> Договір; мастер-колонка (name=ТСКонтрагент) -> Контрагент
# мастер-колонка
t = re.sub(r'(<InputField name="ТСКонтрагент"[^>]*>.*?)<v8:content>Склад</v8:content>',
           r'\1<v8:content>Контрагент</v8:content>', t, count=1, flags=re.S)
# объектный фильтр id=33
t = re.sub(r'(<InputField name="Склад" id="33">.*?)<v8:content>Склад</v8:content>',
           r'\1<v8:content>Договір</v8:content>', t, count=1, flags=re.S)
# прочие заголовки
t = t.replace("<v8:content>Позицій</v8:content>", "<v8:content>Договорів</v8:content>")
t = t.replace("<v8:content>К-ть ЕРП</v8:content>", "<v8:content>Сума ЕРП</v8:content>")
t = t.replace("<v8:content>К-ть Бух</v8:content>", "<v8:content>Сума Бух</v8:content>")
t = t.replace("<v8:content>Розбіжності по складах</v8:content>", "<v8:content>Розбіжності по контрагентах</v8:content>")
t = t.replace("<v8:content>Порівняти залишки товарів ERP та BuhBud</v8:content>",
              "<v8:content>Порівняти залишки взаєморозрахунків ERP та BuhBud</v8:content>")
t = t.replace("<v8:content>Аналіз поточного складу</v8:content>", "<v8:content>Аналіз поточного контрагента</v8:content>")
t = t.replace("<v8:content>Проаналізувати документи всіх позицій поточного складу</v8:content>",
              "<v8:content>Проаналізувати документи всіх позицій поточного контрагента</v8:content>")

# --- новая колонка Договір в таблице документов (клон ТДКонтрагент, новые id) ---
m = re.search(r'(\t*)<InputField name="ТДКонтрагент" id="100">.*?</InputField>\n', t, flags=re.S)
assert m, "блок ТДКонтрагент не найден"
indent = m.group(1)
dog_col = (indent + '<InputField name="ТДДоговор" id="500">\n'
           + indent + '\t<DataPath>Объект.ТаблицаДокументов.Договор</DataPath>\n'
           + indent + '\t<Title>\n'
           + indent + '\t\t<v8:item>\n'
           + indent + '\t\t\t<v8:lang>ru</v8:lang>\n'
           + indent + '\t\t\t<v8:content>Договір</v8:content>\n'
           + indent + '\t\t</v8:item>\n'
           + indent + '\t</Title>\n'
           + indent + '\t<ContextMenu name="ТДДоговорКонтекстноеМеню" id="501"/>\n'
           + indent + '\t<ExtendedTooltip name="ТДДоговорExtendedTooltip" id="502"/>\n'
           + indent + '</InputField>\n')
t = t[:m.end()] + dog_col + t[m.end():]

# --- проверки ---
assert "СинхронизироватьТоварыТолькоТовары" not in t
assert "ТаблицаСкладов" not in t
assert "Объект.ТаблицаДокументов.Договор</DataPath>" in t
assert "АнализТекущегоСчета" not in t
# новые id уникальны среди ВСЕХ id формы (id-пространства категорий разные, но 500+ нигде не заняты)
for nid in ("500", "501", "502"):
    assert t.count('id="%s"' % nid) == 1, "id %s не уникален" % nid

with io.open(PATH, "w", encoding="utf-8", newline="\n") as f:
    f.write(t)
print("OK Form.xml transform")
