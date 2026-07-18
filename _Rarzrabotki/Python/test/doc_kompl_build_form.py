# -*- coding: utf-8 -*-
"""Transform-порт формы обработки СозданиеКомплектацийБух -> Документ.РасчетКомплектаций.ФормаДокумента.
Правки: тип главного реквизита, шапка документа, 2 новые страницы списания, кнопка/команда Акт,
4 форм-реквизита итогов, OnOpen; Module.bsl: подтверждение пересчёта, автоподбор, пересчёт сумм, Акт."""
import os, re, sys, uuid
import xml.dom.minidom

sys.stdout.reconfigure(encoding='utf-8')

OBR_FORM = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Создания комплектаций для бух учета\Forms\Форма\Ext"
CFG = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh"
DST_DIR = CFG + r"\Documents\РасчетКомплектаций\Forms\ФормаДокумента"
DOCXML = CFG + r"\Documents\РасчетКомплектаций.xml"


def read(path):
    raw = open(path, 'rb').read()
    return raw.decode('utf-8-sig').replace('\r\n', '\n')


def write(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    open(path, 'wb').write(b'\xef\xbb\xbf' + text.replace('\n', '\r\n').encode('utf-8'))


# ============================== Form.xml ==============================
text = read(os.path.join(OBR_FORM, "Form.xml"))

# id-контроль исходника
ids0 = [int(m) for m in re.findall(r'id="(-?\d+)"', text)]
assert max(ids0) < 400, f"max id исходника {max(ids0)} >= 400"

# A6. Тип главного реквизита
assert text.count("cfg:ExternalDataProcessorObject.СозданиеКомплектацийБух") == 1
text = text.replace("cfg:ExternalDataProcessorObject.СозданиеКомплектацийБух",
                    "cfg:DocumentObject.РасчетКомплектаций")

# A7. Убрать Save-блок (данные хранит сам документ)
i = text.index("\t\t\t<Save>")
j = text.index("\t\t\t</Save>\n") + len("\t\t\t</Save>\n")
text = text[:i] + text[j:]

# A8. Убрать Title формы (авто-заголовок документа)
k = text.index("<v8:content>Створення комплектацій для бух обліку</v8:content>")
i = text.rindex("\t<Title>", 0, k)
j = text.index("\t</Title>\n", k) + len("\t</Title>\n")
text = text[:i] + text[j:]

# A9. Убрать AutoTitle=false
text = text.replace("\t<AutoTitle>false</AutoTitle>\n", "", 1)

# A10. OnOpen
anchor = "\t\t<Event name=\"OnCreateAtServer\">ПриСозданииНаСервере</Event>\n"
assert anchor in text
text = text.replace(anchor, anchor + "\t\t<Event name=\"OnOpen\">ПриОткрытии</Event>\n", 1)


def title_block(b, title):
    return (f"{b}<Title>\n{b}\t<v8:item>\n{b}\t\t<v8:lang>uk</v8:lang>\n"
            f"{b}\t\t<v8:content>{title}</v8:content>\n{b}\t</v8:item>\n{b}</Title>\n")


def input_field(b, nm, dp, cm, et, title=None, readonly=False, on_change=None, title_top=False):
    s = f"{b}<InputField name=\"{nm}\" id=\"{{{nm}}}\">\n"
    s += f"{b}\t<DataPath>{dp}</DataPath>\n"
    if readonly:
        s += f"{b}\t<ReadOnly>true</ReadOnly>\n"
    if title:
        s += title_block(b + "\t", title)
    if title_top:
        s += f"{b}\t<TitleLocation>Top</TitleLocation>\n"
    s += f"{b}\t<ContextMenu name=\"{nm}КонтекстноеМеню\" id=\"{{{nm}_cm}}\"/>\n"
    s += f"{b}\t<ExtendedTooltip name=\"{nm}РасширеннаяПодсказка\" id=\"{{{nm}_et}}\"/>\n"
    if on_change:
        s += f"{b}\t<Events>\n{b}\t\t<Event name=\"OnChange\">{on_change}</Event>\n{b}\t</Events>\n"
    s += f"{b}</InputField>\n"
    return s


def table_block(b, nm, dp, height, readonly, columns, events=None):
    s = f"{b}<Table name=\"{nm}\" id=\"{{{nm}}}\">\n"
    if readonly:
        s += f"{b}\t<ReadOnly>true</ReadOnly>\n"
    s += f"{b}\t<HeightInTableRows>{height}</HeightInTableRows>\n"
    s += f"{b}\t<DataPath>{dp}</DataPath>\n"
    s += f"{b}\t<RowFilter xsi:nil=\"true\"/>\n"
    s += f"{b}\t<ContextMenu name=\"{nm}КонтекстноеМеню\" id=\"{{{nm}_cm}}\"/>\n"
    s += f"{b}\t<AutoCommandBar name=\"{nm}КоманднаяПанель\" id=\"{{{nm}_acb}}\"/>\n"
    s += f"{b}\t<ExtendedTooltip name=\"{nm}РасширеннаяПодсказка\" id=\"{{{nm}_et}}\"/>\n"
    for add_nm, add_type, suffix in ((f"{nm}СтрокаПоиска", "SearchStringRepresentation", "ssa"),
                                     (f"{nm}СостояниеПросмотра", "ViewStatusRepresentation", "vsa"),
                                     (f"{nm}УправлениеПоиском", "SearchControl", "sca")):
        tag = {"ssa": "SearchStringAddition", "vsa": "ViewStatusAddition", "sca": "SearchControlAddition"}[suffix]
        s += (f"{b}\t<{tag} name=\"{add_nm}\" id=\"{{{nm}_{suffix}}}\">\n"
              f"{b}\t\t<AdditionSource>\n{b}\t\t\t<Item>{nm}</Item>\n{b}\t\t\t<Type>{add_type}</Type>\n{b}\t\t</AdditionSource>\n"
              f"{b}\t\t<ContextMenu name=\"{add_nm}КонтекстноеМеню\" id=\"{{{nm}_{suffix}_cm}}\"/>\n"
              f"{b}\t\t<ExtendedTooltip name=\"{add_nm}РасширеннаяПодсказка\" id=\"{{{nm}_{suffix}_et}}\"/>\n"
              f"{b}\t</{tag}>\n")
    if events:
        s += f"{b}\t<Events>\n"
        for ev_name, handler in events:
            s += f"{b}\t\t<Event name=\"{ev_name}\">{handler}</Event>\n"
        s += f"{b}\t</Events>\n"
    s += f"{b}\t<ChildItems>\n"
    s += columns
    s += f"{b}\t</ChildItems>\n{b}</Table>\n"
    return s


def group_block(b, nm, title, children, strong=False):
    s = f"{b}<UsualGroup name=\"{nm}\" id=\"{{{nm}}}\">\n"
    s += title_block(b + "\t", title)
    s += f"{b}\t<Group>Horizontal</Group>\n{b}\t<Behavior>Usual</Behavior>\n"
    if strong:
        s += f"{b}\t<Representation>StrongSeparation</Representation>\n"
    s += f"{b}\t<ExtendedTooltip name=\"{nm}РасширеннаяПодсказка\" id=\"{{{nm}_et}}\"/>\n"
    s += f"{b}\t<ChildItems>\n{children}{b}\t</ChildItems>\n{b}</UsualGroup>\n"
    return s


def page_block(b, nm, title, children):
    s = f"{b}<Page name=\"{nm}\" id=\"{{{nm}}}\">\n"
    s += title_block(b + "\t", title)
    s += f"{b}\t<ExtendedTooltip name=\"{nm}РасширеннаяПодсказка\" id=\"{{{nm}_et}}\"/>\n"
    s += f"{b}\t<ChildItems>\n{children}{b}\t</ChildItems>\n{b}</Page>\n"
    return s


# A1. Шапка документа перед основной командной панелью
cb_anchor = "\t\t<CommandBar name=\"ОсновнаяКоманднаяПанель\" id=\"1\">"
assert cb_anchor in text
b = "\t\t"
fields = ""
fields += input_field(b + "\t\t", "Номер", "Объект.Номер", None, None, readonly=True)
fields += input_field(b + "\t\t", "ДатаДокумента", "Объект.Дата", None, None)
fields += input_field(b + "\t\t", "СтатусПоле", "Объект.Статус", None, None, readonly=True)
fields += input_field(b + "\t\t", "ОтветственныйПоле", "Объект.Ответственный", None, None)
fields += input_field(b + "\t\t", "КомментарийПоле", "Объект.Комментарий", None, None)
header_group = group_block(b, "ГруппаШапка", "Документ", fields)
text = text.replace(cb_anchor, header_group + cb_anchor, 1)

# A2. Все печати — в выпадающее меню «Друк» (иначе командная панель переполняется,
# хвост кнопок уезжает в «Ещё»). Существующие 5 кнопок печати вырезаются и
# вкладываются в Popup, туда же добавляются Акт и Розбіжності.
start_print = text.index("\t\t\t\t<Button name=\"КнопкаПечать\" id=\"160\">")
end_marker = ("<ExtendedTooltip name=\"КнопкаПечатьАнализССОднаРасширеннаяПодсказка\" id=\"282\"/>\n"
              "\t\t\t\t</Button>\n")
end_print = text.index(end_marker) + len(end_marker)
print_block = text[start_print:end_print]
print_block = "".join(("\t" + line if line.strip() else line)
                      for line in print_block.splitlines(keepends=True))
akt_button = ("\t\t\t\t\t\t<Button name=\"КнопкаПечатьАкт\" id=\"{КнопкаПечатьАкт}\">\n"
              "\t\t\t\t\t\t\t<Type>UsualButton</Type>\n"
              "\t\t\t\t\t\t\t<CommandName>Form.Command.ПечатьАкт</CommandName>\n"
              + title_block("\t\t\t\t\t\t\t", "Акт понаднормового списання") +
              "\t\t\t\t\t\t\t<ExtendedTooltip name=\"КнопкаПечатьАктРасширеннаяПодсказка\" id=\"{КнопкаПечатьАкт_et}\"/>\n"
              "\t\t\t\t\t\t</Button>\n")
rozb_button = ("\t\t\t\t\t\t<Button name=\"КнопкаПечатьРозбижности\" id=\"{КнопкаПечатьРозбижности}\">\n"
               "\t\t\t\t\t\t\t<Type>UsualButton</Type>\n"
               "\t\t\t\t\t\t\t<CommandName>Form.Command.ПечатьРозбижности</CommandName>\n"
               + title_block("\t\t\t\t\t\t\t", "Друк розбіжностей одиниць") +
               "\t\t\t\t\t\t\t<ExtendedTooltip name=\"КнопкаПечатьРозбижностиРасширеннаяПодсказка\" id=\"{КнопкаПечатьРозбижности_et}\"/>\n"
               "\t\t\t\t\t\t</Button>\n")
popup_print = ("\t\t\t\t<Popup name=\"ПодменюПечать\" id=\"{ПодменюПечать}\">\n"
               + title_block("\t\t\t\t\t", "Друк")
               + "\t\t\t\t\t<ExtendedTooltip name=\"ПодменюПечатьРасширеннаяПодсказка\" id=\"{ПодменюПечать_et}\"/>\n"
               + "\t\t\t\t\t<ChildItems>\n"
               + print_block + akt_button + rozb_button
               + "\t\t\t\t\t</ChildItems>\n"
               + "\t\t\t\t</Popup>\n")
text = text[:start_print] + popup_print + text[end_print:]
# внутри Popup кнопка «Друк» остаётся с иконкой и коротким названием — это пункт меню

# A2b. Колонка-флажок «Розбіжність од.» в таблице анализа (после ТАЕкономіяСума)
ekon_dp = "<DataPath>ТаблицаАнализа.ЕкономіяСума</DataPath>"
i2b = text.index(ekon_dp)
j2b = text.index("</InputField>\n", i2b) + len("</InputField>\n")
checkbox = ("\t\t\t\t\t\t\t\t<CheckBoxField name=\"ТАРасхождениеЕдиниц\" id=\"{ТАРасхождениеЕдиниц}\">\n"
            "\t\t\t\t\t\t\t\t\t<DataPath>ТаблицаАнализа.РасхождениеЕдиниц</DataPath>\n"
            + title_block("\t\t\t\t\t\t\t\t\t", "Розбіжність од.") +
            "\t\t\t\t\t\t\t\t\t<ContextMenu name=\"ТАРасхождениеЕдиницКонтекстноеМеню\" id=\"{ТАРасхождениеЕдиниц_cm}\"/>\n"
            "\t\t\t\t\t\t\t\t\t<ExtendedTooltip name=\"ТАРасхождениеЕдиницРасширеннаяПодсказка\" id=\"{ТАРасхождениеЕдиниц_et}\"/>\n"
            "\t\t\t\t\t\t\t\t</CheckBoxField>\n")
text = text[:j2b] + checkbox + text[j2b:]

# A2c. Колонка VT ТаблицаАнализа.РасхождениеЕдиниц (Boolean)
cols_anchor = "\t\t</Columns>"
assert text.count(cols_anchor) == 1
vt_col = ("\t\t\t<Column name=\"РасхождениеЕдиниц\" id=\"{col_РасхождениеЕдиниц}\">\n"
          "\t\t\t\t<Type>\n\t\t\t\t\t<v8:Type>xs:boolean</v8:Type>\n\t\t\t\t</Type>\n"
          "\t\t\t</Column>\n")
text = text.replace(cols_anchor, vt_col + cols_anchor, 1)

# A3. Две страницы списания перед СтраницаДокументы
# Глубины: Page=4 таба, Table/группа на странице=6, колонки/поля в них=8
page_anchor = "\t\t\t\t<Page name=\"СтраницаДокументы\" id=\"83\">"
assert page_anchor in text
pb, tb, cb2 = "\t\t\t\t", "\t\t\t\t\t\t", "\t\t\t\t\t\t\t"

cols_norm = ""
for col in ("Номенклатура", "ОбщееНазвание", "Склад", "Счет", "Единица", "Количество", "Цена", "Сумма"):
    cols_norm += input_field(cb2 + "\t", "СПН" + col, "Объект.СписаниеПоНормам." + col, None, None)
itogi_norm = (input_field(cb2 + "\t", "ИтогоНормКоличествоПоле", "ИтогоНормКоличество", None, None,
                          title="Кількість", readonly=True, title_top=True)
              + input_field(cb2 + "\t", "ИтогоНормСуммаПоле", "ИтогоНормСумма", None, None,
                            title="Сума (без ПДВ)", readonly=True, title_top=True))
page_norm = page_block(pb, "СтраницаПоНормам", "Списання за нормами",
                       table_block(tb, "СписаниеПоНормам", "Объект.СписаниеПоНормам", 12, True, cols_norm)
                       + group_block(tb, "ГруппаИтогиПоНормам", "Підсумки", itogi_norm, strong=True))

cols_over = ""
cols_over += input_field(cb2 + "\t", "ССННоменклатура", "Объект.СписаниеСверхНормы.Номенклатура", None, None,
                         on_change="СверхНормыНоменклатураПриИзменении")
cols_over += input_field(cb2 + "\t", "ССНОбщееНазвание", "Объект.СписаниеСверхНормы.ОбщееНазвание", None, None,
                         readonly=True)
cols_over += input_field(cb2 + "\t", "ССНСклад", "Объект.СписаниеСверхНормы.Склад", None, None,
                         on_change="СверхНормыСкладПриИзменении")
cols_over += input_field(cb2 + "\t", "ССНСчет", "Объект.СписаниеСверхНормы.Счет", None, None)
cols_over += input_field(cb2 + "\t", "ССНЕдиница", "Объект.СписаниеСверхНормы.Единица", None, None)
cols_over += input_field(cb2 + "\t", "ССНКоличество", "Объект.СписаниеСверхНормы.Количество", None, None,
                         on_change="СверхНормыКоличествоПриИзменении")
cols_over += input_field(cb2 + "\t", "ССНЦена", "Объект.СписаниеСверхНормы.Цена", None, None,
                         on_change="СверхНормыЦенаПриИзменении")
cols_over += input_field(cb2 + "\t", "ССНСумма", "Объект.СписаниеСверхНормы.Сумма", None, None, readonly=True)
cols_over += input_field(cb2 + "\t", "ССНПричина", "Объект.СписаниеСверхНормы.Причина", None, None)
itogi_over = (input_field(cb2 + "\t", "ИтогоСверхКоличествоПоле", "ИтогоСверхКоличество", None, None,
                          title="Кількість", readonly=True, title_top=True)
              + input_field(cb2 + "\t", "ИтогоСверхСуммаПоле", "ИтогоСверхСумма", None, None,
                            title="Сума (без ПДВ)", readonly=True, title_top=True))
page_over = page_block(pb, "СтраницаПонадНорму", "Списання понад норму",
                       table_block(tb, "СписаниеСверхНормы", "Объект.СписаниеСверхНормы", 12, False, cols_over,
                                   events=[("AfterDeleteRow", "СверхНормыПослеУдаления")])
                       + group_block(tb, "ГруппаИтогиСверхНормы", "Підсумки", itogi_over, strong=True))

text = text.replace(page_anchor, page_norm + page_over + page_anchor, 1)

# A4. Форм-реквизиты итогов
attrs_anchor = "\t</Attributes>"
assert attrs_anchor in text
DEC = ("\t\t<Attribute name=\"{nm}\" id=\"{{attr_{nm}}}\">\n"
       "\t\t\t<Type>\n\t\t\t\t<v8:Type>xs:decimal</v8:Type>\n"
       "\t\t\t\t<v8:NumberQualifiers>\n\t\t\t\t\t<v8:Digits>15</v8:Digits>\n"
       "\t\t\t\t\t<v8:FractionDigits>{fd}</v8:FractionDigits>\n"
       "\t\t\t\t\t<v8:AllowedSign>Any</v8:AllowedSign>\n\t\t\t\t</v8:NumberQualifiers>\n\t\t\t</Type>\n"
       "\t\t</Attribute>\n")
new_attrs = (DEC.format(nm="ИтогоНормКоличество", fd=3) + DEC.format(nm="ИтогоНормСумма", fd=2)
             + DEC.format(nm="ИтогоСверхКоличество", fd=3) + DEC.format(nm="ИтогоСверхСумма", fd=2))
text = text.replace(attrs_anchor, new_attrs + attrs_anchor, 1)

# A5. Команда ПечатьАкт
cmd_anchor = "\t</Commands>"
assert cmd_anchor in text
akt_cmd = ("\t\t<Command name=\"ПечатьАкт\" id=\"{cmd_ПечатьАкт}\">\n"
           + title_block("\t\t\t", "Акт понаднормового списання") +
           "\t\t\t<Action>ПечатьАкт</Action>\n\t\t</Command>\n")
rozb_cmd = ("\t\t<Command name=\"ПечатьРозбижности\" id=\"{cmd_ПечатьРозбижности}\">\n"
            + title_block("\t\t\t", "Друк розбіжностей одиниць") +
            "\t\t\t<Action>ПечатьРозбижности</Action>\n\t\t</Command>\n")
text = text.replace(cmd_anchor, akt_cmd + rozb_cmd + cmd_anchor, 1)

# Раздача id всем {placeholder}
placeholders = []
for m in re.finditer(r'\{([А-Яа-яЁёA-Za-z_іІїЇєЄ0-9]+)\}', text):
    if m.group(1) not in placeholders:
        placeholders.append(m.group(1))
next_id = 400
mapping = {}
for ph in placeholders:
    mapping[ph] = next_id
    next_id += 1
text = re.sub(r'\{([А-Яа-яЁёA-Za-z_іІїЇєЄ0-9]+)\}', lambda m: str(mapping[m.group(1)]), text)

# Контроли
ids = re.findall(r'id="(-?\d+)"', text)
assert len(ids) == len(set(ids)), "дубли id в Form.xml!"
assert "ExternalDataProcessorObject" not in text
write(os.path.join(DST_DIR, "Ext", "Form.xml"), text)
xml.dom.minidom.parse(os.path.join(DST_DIR, "Ext", "Form.xml"))
print(f"Form.xml OK: элементов id={len(ids)}, новых={len(placeholders)}, max id={max(int(x) for x in ids)}")

# ============================== Module.bsl ==============================
mod = read(os.path.join(OBR_FORM, "Form", "Module.bsl"))

# 1. ПриСозданииНаСервере -> новая версия (дефолты теперь в ОбработкаЗаполнения объекта)
i = mod.index("&НаСервере\nПроцедура ПриСозданииНаСервере")
j = mod.index("КонецПроцедуры", i) + len("КонецПроцедуры")
NEW_ONCREATE = '''&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	// Метрик-панель «Підсумки»: підсвічування «В нормі» (зелений) і «Понад» (янтарний).
	Элементы.ИтогоВНорме.ЦветФона = Новый Цвет(214, 240, 214);
	Элементы.ИтогоПонад.ЦветФона  = Новый Цвет(250, 235, 200);
	// Підсвічування рядків «понад норму» в таблиці аналізу.
	УстановитьУсловноеОформлениеСтрок();
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	// Відновлення робочого місця зі збереженого розрахунку.
	ПрименитьОтборИИтоги();
	ПересчитатьИтогиСписания();
КонецПроцедуры'''
mod = mod[:i] + NEW_ONCREATE + mod[j:]

# 2. Рассчитать -> подтверждение при статусе «Документи створено» + запись после расчёта
i = mod.index("&НаКлиенте\nПроцедура Рассчитать(Команда)")
j = mod.index("КонецПроцедуры", i) + len("КонецПроцедуры")
NEW_RASSCHITAT = '''&НаКлиенте
Процедура Рассчитать(Команда)
	Если Объект.Статус = ПредопределенноеЗначение("Перечисление.СтатусыРасчетаКомплектаций.ДокументыСозданы") Тогда
		Оповещение = Новый ОписаниеОповещения("РассчитатьПродолжение", ЭтотОбъект);
		ПоказатьВопрос(Оповещение,
			"Документи вже створено за цим розрахунком. Перезаповнити розрахунок?",
			РежимДиалогаВопрос.ДаНет);
	Иначе
		РассчитатьПродолжение(КодВозвратаДиалога.Да, Неопределено);
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура РассчитатьПродолжение(РезультатВопроса, ДопПараметры) Экспорт
	Если РезультатВопроса <> КодВозвратаДиалога.Да Тогда
		Возврат;
	КонецЕсли;
	РассчитатьНаСервере();
	ПрименитьОтборИИтоги();
	ПересчитатьИтогиСписания();
	Записать();
КонецПроцедуры'''
mod = mod[:i] + NEW_RASSCHITAT + mod[j:]

# 3. Запись формы после заполнения документов (3 режима)
for args in ("Истина, Истина", "Истина, Ложь", "Ложь, Истина"):
    old = f"\tЗаполнитьДокументыНаСервере({args});\nКонецПроцедуры"
    assert old in mod, args
    mod = mod.replace(old, f"\tЗаполнитьДокументыНаСервере({args});\n\tЗаписать();\nКонецПроцедуры", 1)

# 3b. ПрименитьОтборИИтоги: перенос колонки РасхождениеЕдиниц
anchor3b = "\t\tНС.Этап            = Стр.Этап;\n"
assert anchor3b in mod
mod = mod.replace(anchor3b, anchor3b + "\t\tНС.РасхождениеЕдиниц = Стр.РасхождениеЕдиниц;\n", 1)

# 3c. УстановитьУсловноеОформлениеСтрок: второе правило — розовый по расхождению единиц
anchor3c = "\tЭлемент.Оформление.УстановитьЗначениеПараметра(\"ЦветФона\", Новый Цвет(250, 235, 200));\n\nКонецПроцедуры"
assert anchor3c in mod
UO_ROZB = '''\tЭлемент.Оформление.УстановитьЗначениеПараметра("ЦветФона", Новый Цвет(250, 235, 200));

\t// Розбіжність одиниць — рожевим (правило пізніше — перекриває янтарне понаднорми).
\tЭлемент = УсловноеОформление.Элементы.Добавить();
\tЭлемент.Представление = "Розбіжність одиниць";

\tПолеЭлемента = Элемент.Поля.Элементы.Добавить();
\tПолеЭлемента.Поле = Новый ПолеКомпоновкиДанных(Элементы.ТаблицаАнализа.Имя);

\tОтборЭлемента = Элемент.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
\tОтборЭлемента.ЛевоеЗначение = Новый ПолеКомпоновкиДанных("ТаблицаАнализа.РасхождениеЕдиниц");
\tОтборЭлемента.ВидСравнения = ВидСравненияКомпоновкиДанных.Равно;
\tОтборЭлемента.ПравоеЗначение = Истина;

\tЭлемент.Оформление.УстановитьЗначениеПараметра("ЦветФона", Новый Цвет(250, 215, 215));

КонецПроцедуры'''
mod = mod.replace(anchor3c, UO_ROZB, 1)

# 3d. РассчитатьНаСервере: сообщение о количестве расхождений
anchor3d = "\tСообщить(\"Розраховано позицій: \" + Объект.ТабличнаяЧастьОстатков.Количество());\n"
assert anchor3d in mod
mod = mod.replace(anchor3d, anchor3d
                  + "\tРозбіжностей = 0;\n"
                  + "\tДля Каждого Стр Из Объект.ТабличнаяЧастьОстатков Цикл\n"
                  + "\t\tЕсли Стр.РасхождениеЕдиниц Тогда\n"
                  + "\t\t\tРозбіжностей = Розбіжностей + 1;\n"
                  + "\t\tКонецЕсли;\n"
                  + "\tКонецЦикла;\n"
                  + "\tЕсли Розбіжностей > 0 Тогда\n"
                  + "\t\tСообщить(\"Розбіжностей одиниць: \" + Розбіжностей\n"
                  + "\t\t\t+ \" рядків (рожеві рядки; див. «Друк розбіжностей одиниць»).\");\n"
                  + "\tКонецЕсли;\n", 1)

# 4. Новые обработчики в конец модуля
NEW_HANDLERS = '''

&НаКлиенте
Процедура СверхНормыКоличествоПриИзменении(Элемент)
	Стр = Элементы.СписаниеСверхНормы.ТекущиеДанные;
	Если Стр = Неопределено Тогда
		Возврат;
	КонецЕсли;
	Стр.Сумма = Окр(Стр.Количество * Стр.Цена, 2);
	Если ЗначениеЗаполнено(Стр.Номенклатура) И ЗначениеЗаполнено(Стр.Склад) Тогда
		ПроверитьОстатокСтрокиНаСервере(Стр.Номенклатура, Стр.Склад, Стр.Количество);
	КонецЕсли;
	ПересчитатьИтогиСписания();
КонецПроцедуры

&НаКлиенте
Процедура СверхНормыПослеУдаления(Элемент)
	ПересчитатьИтогиСписания();
КонецПроцедуры

&НаКлиенте
Процедура СверхНормыЦенаПриИзменении(Элемент)
	Стр = Элементы.СписаниеСверхНормы.ТекущиеДанные;
	Если Стр = Неопределено Тогда
		Возврат;
	КонецЕсли;
	Стр.Сумма = Окр(Стр.Количество * Стр.Цена, 2);
	ПересчитатьИтогиСписания();
КонецПроцедуры

&НаКлиенте
Процедура СверхНормыНоменклатураПриИзменении(Элемент)
	ПодобратьДанныеСтрокиСверхНормы();
КонецПроцедуры

&НаКлиенте
Процедура СверхНормыСкладПриИзменении(Элемент)
	ПодобратьДанныеСтрокиСверхНормы();
КонецПроцедуры

&НаКлиенте
Процедура ПодобратьДанныеСтрокиСверхНормы()
	Стр = Элементы.СписаниеСверхНормы.ТекущиеДанные;
	Если Стр = Неопределено ИЛИ НЕ ЗначениеЗаполнено(Стр.Номенклатура) ИЛИ НЕ ЗначениеЗаполнено(Стр.Склад) Тогда
		Возврат;
	КонецЕсли;
	Данные = ДанныеОстаткаНаСервере(Стр.Номенклатура, Стр.Склад);
	Если Данные = Неопределено Тогда
		Сообщить("Залишок за цією номенклатурою на складі не знайдено (рахунок/ціну не підібрано).");
		Возврат;
	КонецЕсли;
	Стр.Счет = Данные.Счет;
	Стр.Единица = Данные.Единица;
	Стр.ОбщееНазвание = Данные.ОбщееНазвание;
	Стр.Цена = Данные.Цена;
	Стр.Сумма = Окр(Стр.Количество * Стр.Цена, 2);
	ПересчитатьИтогиСписания();
КонецПроцедуры

&НаСервере
Функция ДанныеОстаткаНаСервере(Номенклатура, Склад)
	ОбъектДокумента = РеквизитФормыВЗначение("Объект");
	Возврат ОбъектДокумента.ДанныеОстаткаДляСтроки(Номенклатура, Склад);
КонецФункции

&НаСервере
Процедура ПроверитьОстатокСтрокиНаСервере(Номенклатура, Склад, Количество)
	ОбъектДокумента = РеквизитФормыВЗначение("Объект");
	Данные = ОбъектДокумента.ДанныеОстаткаДляСтроки(Номенклатура, Склад);
	Если Данные = Неопределено Тогда
		Сообщить("Увага: залишку немає — списуємо без покриття залишком.");
	ИначеЕсли Количество > Данные.Остаток Тогда
		Сообщить("Увага: кількість " + Количество + " перевищує залишок " + Данные.Остаток + ".");
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ПересчитатьИтогиСписания()
	ИтогоНормКоличество = 0;
	ИтогоНормСумма = 0;
	Для Каждого Стр Из Объект.СписаниеПоНормам Цикл
		ИтогоНормКоличество = ИтогоНормКоличество + Стр.Количество;
		ИтогоНормСумма = ИтогоНормСумма + Стр.Сумма;
	КонецЦикла;
	ИтогоСверхКоличество = 0;
	ИтогоСверхСумма = 0;
	Для Каждого Стр Из Объект.СписаниеСверхНормы Цикл
		ИтогоСверхКоличество = ИтогоСверхКоличество + Стр.Количество;
		ИтогоСверхСумма = ИтогоСверхСумма + Стр.Сумма;
	КонецЦикла;
КонецПроцедуры

&НаКлиенте
Процедура ПечатьАкт(Команда)
	ТабДок = ПечатьАктНаСервере();
	Если ТабДок <> Неопределено Тогда
		ТабДок.Показать("Акт на списання понаднормових матеріалів");
	КонецЕсли;
КонецПроцедуры

&НаСервере
Функция ПечатьАктНаСервере()
	Если Объект.СписаниеСверхНормы.Количество() = 0 Тогда
		Сообщить("Таблиця «Списання понад норму» порожня.");
		Возврат Неопределено;
	КонецЕсли;
	ОбъектДокумента = РеквизитФормыВЗначение("Объект");
	Возврат ОбъектДокумента.СформироватьПечатьАкт();
КонецФункции

&НаКлиенте
Процедура ПечатьРозбижности(Команда)
	ТабДок = ПечатьРозбижностиНаСервере();
	Если ТабДок <> Неопределено Тогда
		ТабДок.Показать("Розбіжності одиниць (норми - фактичний прихід)");
	КонецЕсли;
КонецПроцедуры

&НаСервере
Функция ПечатьРозбижностиНаСервере()
	Если Объект.ТабличнаяЧастьОстатков.Количество() = 0 Тогда
		Сообщить("Спочатку натисніть «Розрахувати».");
		Возврат Неопределено;
	КонецЕсли;
	ОбъектДокумента = РеквизитФормыВЗначение("Объект");
	Возврат ОбъектДокумента.СформироватьПечатьРозбіжностіОдиниць();
КонецФункции
'''
mod = mod.rstrip('\n') + '\n' + NEW_HANDLERS

n_proc = len(re.findall(r"^(?:Процедура|Функция) ", mod, re.M))
n_end = len(re.findall(r"^Конец(?:Процедуры|Функции)", mod, re.M))
assert n_proc == n_end, f"непарность: {n_proc}/{n_end}"
write(os.path.join(DST_DIR, "Ext", "Form", "Module.bsl"), mod)
print(f"Module.bsl OK: процедур={n_proc}")

# ============================== Обёртка формы ==============================
# UUID обёртки СТАБИЛЬНЫЙ: переиспользуем из существующего файла, иначе частичная
# перезагрузка видит «новую форму без родителя» и падает.
wrapper_path = CFG + r"\Documents\РасчетКомплектаций\Forms\ФормаДокумента.xml"
if os.path.exists(wrapper_path):
    _m = re.search(r'<Form uuid="([0-9a-f-]{36})"', read(wrapper_path))
    form_uuid = _m.group(1)
else:
    form_uuid = str(uuid.uuid4())
WRAPPER = '''<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses" xmlns:app="http://v8.1c.ru/8.2/managed-application/core" xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" xmlns:cmi="http://v8.1c.ru/8.2/managed-application/cmi" xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" xmlns:style="http://v8.1c.ru/8.1/data/ui/style" xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" xmlns:xen="http://v8.1c.ru/8.3/xcf/enums" xmlns:xpr="http://v8.1c.ru/8.3/xcf/predef" xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13">
	<Form uuid="''' + form_uuid + '''">
		<Properties>
			<Name>ФормаДокумента</Name>
			<Synonym>
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Форма документа</v8:content>
				</v8:item>
				<v8:item>
					<v8:lang>uk</v8:lang>
					<v8:content>Форма документа</v8:content>
				</v8:item>
			</Synonym>
			<Comment/>
			<FormType>Managed</FormType>
			<IncludeHelpInContents>false</IncludeHelpInContents>
			<UsePurposes>
				<v8:Value xsi:type="app:ApplicationUsePurpose">PlatformApplication</v8:Value>
				<v8:Value xsi:type="app:ApplicationUsePurpose">MobilePlatformApplication</v8:Value>
			</UsePurposes>
		</Properties>
	</Form>
</MetaDataObject>
'''
write(CFG + r"\Documents\РасчетКомплектаций\Forms\ФормаДокумента.xml", WRAPPER)
xml.dom.minidom.parse(CFG + r"\Documents\РасчетКомплектаций\Forms\ФормаДокумента.xml")
print("Обёртка формы OK")

# ============================== Регистрация в xml документа ==============================
raw = open(DOCXML, 'rb').read()
doc = raw.decode('utf-8-sig').replace('\r\n', '\n')
if "<Form>ФормаДокумента</Form>" not in doc:
    k = doc.index("\t\t\t<TabularSection")
    doc = doc[:k] + "\t\t\t<Form>ФормаДокумента</Form>\n" + doc[k:]
    open(DOCXML, 'wb').write(b'\xef\xbb\xbf' + doc.replace('\n', '\r\n').encode('utf-8'))
    xml.dom.minidom.parse(DOCXML)
    print("Регистрация <Form> OK")
else:
    print("Регистрация <Form> уже была")
print("FORM BUILD PASS")
