# -*- coding: utf-8 -*-
"""Генератор керованої форми обробки «Друк М-29».

Пише `_Rarzrabotki/Обработки/Печать М29/Forms/Форма/Ext/Form.xml` та обгортку
`Forms/Форма.xml` (uuid обгортки СТАБІЛЬНИЙ — переюзається з наявного файлу).

Компоновка за каноном bas-form-layout: «один гумовий елемент» — таблиця Документы;
поля підписів TitleLocation=Top (при Left форма отримує горизонтальну прокрутку).

Запуск: python _Rarzrabotki/Python/test/m29_obr_build_form.py
"""
import io
import os
import re
import sys
import uuid

sys.stdout.reconfigure(encoding='utf-8')

БАЗА = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Печать М29\Forms\Форма"
ФАЙЛ_ФОРМЫ = os.path.join(БАЗА, "Ext", "Form.xml")
ФАЙЛ_ОБЕРТКИ = os.path.join(os.path.dirname(БАЗА), "Форма.xml")
ИМЯ_ОБРАБОТКИ = "ПечатьМ29"

NS_FORM = ('xmlns="http://v8.1c.ru/8.3/xcf/logform" '
           'xmlns:app="http://v8.1c.ru/8.2/managed-application/core" '
           'xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" '
           'xmlns:dcscor="http://v8.1c.ru/8.1/data-composition-system/core" '
           'xmlns:dcsset="http://v8.1c.ru/8.1/data-composition-system/settings" '
           'xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" '
           'xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" '
           'xmlns:style="http://v8.1c.ru/8.1/data/ui/style" '
           'xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" '
           'xmlns:v8="http://v8.1c.ru/8.1/data/core" '
           'xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" '
           'xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" '
           'xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" '
           'xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" '
           'xmlns:xs="http://www.w3.org/2001/XMLSchema" '
           'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13"')

NS_MD = ('xmlns="http://v8.1c.ru/8.3/MDClasses" '
         'xmlns:app="http://v8.1c.ru/8.2/managed-application/core" '
         'xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" '
         'xmlns:cmi="http://v8.1c.ru/8.2/managed-application/cmi" '
         'xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" '
         'xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" '
         'xmlns:style="http://v8.1c.ru/8.1/data/ui/style" '
         'xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" '
         'xmlns:v8="http://v8.1c.ru/8.1/data/core" '
         'xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" '
         'xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" '
         'xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" '
         'xmlns:xen="http://v8.1c.ru/8.3/xcf/enums" '
         'xmlns:xpr="http://v8.1c.ru/8.3/xcf/predef" '
         'xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" '
         'xmlns:xs="http://www.w3.org/2001/XMLSchema" '
         'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13"')

_id = [0]


def nid():
    _id[0] += 1
    return _id[0]


def назва(текст, о):
    return (f"{о}<Title>\n{о}\t<v8:item>\n{о}\t\t<v8:lang>uk</v8:lang>\n"
            f"{о}\t\t<v8:content>{текст}</v8:content>\n{о}\t</v8:item>\n{о}</Title>")


def поле(имя, путь, о, титул=None, сверху=False, только_чтение=False, тип="InputField",
         доп=""):
    т = "\t" * о
    части = [f'{т}<{тип} name="{имя}" id="{nid()}">', f"{т}\t<DataPath>{путь}</DataPath>"]
    if тип == "CheckBoxField":
        части.append(f"{т}\t<CheckBoxType>Auto</CheckBoxType>")
    if только_чтение:
        части.append(f"{т}\t<ReadOnly>true</ReadOnly>")
    if титул:
        части.append(назва(титул, т + "\t"))
    if сверху:
        части.append(f"{т}\t<TitleLocation>Top</TitleLocation>")
    if доп:
        части.append(доп)
    части.append(f'{т}\t<ContextMenu name="{имя}КонтекстноеМеню" id="{nid()}"/>')
    части.append(f'{т}\t<ExtendedTooltip name="{имя}РасширеннаяПодсказка" id="{nid()}"/>')
    части.append(f"{т}</{тип}>")
    return "\n".join(части)


def группа(имя, о, дети, горизонтальная=True, титул=None, равные=False, растягивать=None):
    т = "\t" * о
    части = [f'{т}<UsualGroup name="{имя}" id="{nid()}">']
    if титул:
        части.append(назва(титул, т + "\t"))
    # ⚠ VerticalStretch идёт ДО Group — порядок элементов в XDTO-схеме жёсткий
    if растягивать is not None:
        части.append(f"{т}\t<VerticalStretch>{'true' if растягивать else 'false'}</VerticalStretch>")
    части.append(f"{т}\t<Group>{'Horizontal' if горизонтальная else 'Vertical'}</Group>")
    части.append(f"{т}\t<Behavior>Usual</Behavior>")
    части.append(f"{т}\t<Representation>None</Representation>")
    if равные:
        части.append(f"{т}\t<ChildItemsWidth>Equal</ChildItemsWidth>")
    части.append(f"{т}\t<ShowTitle>{'true' if титул else 'false'}</ShowTitle>")
    части.append(f'{т}\t<ExtendedTooltip name="{имя}РасширеннаяПодсказка" id="{nid()}"/>')
    части.append(f"{т}\t<ChildItems>")
    части.append("\n".join(дети))
    части.append(f"{т}\t</ChildItems>")
    части.append(f"{т}</UsualGroup>")
    return "\n".join(части)


def таблица(имя, путь, о, колонки, титул=None, высота=None):
    т = "\t" * о
    части = [f'{т}<Table name="{имя}" id="{nid()}">']
    if высота:
        части.append(f"{т}\t<HeightInTableRows>{высота}</HeightInTableRows>")
    части.append(f"{т}\t<DataPath>{путь}</DataPath>")
    if титул:
        части.append(назва(титул, т + "\t"))
    части.append(f'{т}\t<RowFilter xsi:nil="true"/>')
    части.append(f'{т}\t<ContextMenu name="{имя}КонтекстноеМеню" id="{nid()}"/>')
    части.append(f'{т}\t<AutoCommandBar name="{имя}КоманднаяПанель" id="{nid()}"/>')
    части.append(f'{т}\t<ExtendedTooltip name="{имя}РасширеннаяПодсказка" id="{nid()}"/>')
    части.append(f"{т}\t<ChildItems>")
    части.append("\n".join(колонки))
    части.append(f"{т}\t</ChildItems>")
    части.append(f"{т}</Table>")
    return "\n".join(части)


def кнопка(имя, команда, титул, о, по_умолчанию=False):
    т = "\t" * о
    части = [f'{т}<Button name="{имя}" id="{nid()}">', f"{т}\t<Type>UsualButton</Type>"]
    if по_умолчанию:
        части.append(f"{т}\t<DefaultButton>true</DefaultButton>")
    части.append(f"{т}\t<CommandName>Form.Command.{команда}</CommandName>")
    части.append(назва(титул, т + "\t"))
    части.append(f'{т}\t<ExtendedTooltip name="{имя}РасширеннаяПодсказка" id="{nid()}"/>')
    части.append(f"{т}</Button>")
    return "\n".join(части)


def страница(имя, титул, о, дети):
    т = "\t" * о
    части = [f'{т}<Page name="{имя}" id="{nid()}">']
    части.append(назва(титул, т + "\t"))
    части.append(f'{т}\t<ExtendedTooltip name="{имя}РасширеннаяПодсказка" id="{nid()}"/>')
    части.append(f"{т}\t<ChildItems>")
    части.append("\n".join(дети))
    части.append(f"{т}\t</ChildItems>")
    части.append(f"{т}</Page>")
    return "\n".join(части)


def main():
    # ---------------- командная панель ----------------
    кнопки = [
        кнопка("КнопкаЗаполнить", "Заполнить", "Заповнити", 2),
        кнопка("КнопкаОтметитьВсе", "ОтметитьВсе", "Відмітити всі", 2),
        кнопка("КнопкаСнятьОтметки", "СнятьОтметки", "Зняти відмітки", 2),
        кнопка("КнопкаПечатьМ29", "ПечатьМ29", "Друк М-29", 2, по_умолчанию=True),
    ]

    # ---------------- страница «Документи» ----------------
    ряд1 = группа("ГруппаПериод", 6, [
        поле("ПериодС", "Объект.ПериодС", 8),
        поле("ПериодПо", "Объект.ПериодПо", 8),
        поле("Организация", "Объект.Организация", 8),
    ], равные=True, растягивать=False)

    ряд2 = группа("ГруппаВиды", 6, [
        поле("ВключатьТребованияНакладные", "Объект.ВключатьТребованияНакладные", 8,
             тип="CheckBoxField", титул="Вимоги-накладні"),
        поле("ВключатьСписанияТоваров", "Объект.ВключатьСписанияТоваров", 8,
             тип="CheckBoxField", титул="Списання товарів"),
        поле("ВключатьКомплектации", "Объект.ВключатьКомплектации", 8,
             тип="CheckBoxField", титул="Комплектації"),
        поле("ВключатьПередачиМалоценки", "Объект.ВключатьПередачиМалоценки", 8,
             тип="CheckBoxField", титул="Передачі малоцінки"),
    ], равные=True, растягивать=False)

    # Два фильтра-списка рядом: канон bas-form-layout — не более 2 таблиц в ряд
    # + ChildItemsWidth=Equal. Обе фиксированной высоты, чтобы не отбирать
    # вертикаль у резиновой таблицы документов.
    ряд3 = группа("ГруппаСпискиОтбора", 6, [
        таблица("Склады", "Объект.Склады", 8,
                [поле("СкладыСклад", "Объект.Склады.Склад", 10)],
                титул="Склади (можна групи)", высота=4),
        таблица("Подразделения", "Объект.Подразделения", 8,
                [поле("ПодразделенияПодразделение", "Объект.Подразделения.Подразделение", 10)],
                титул="Підрозділи (з підлеглими)", высота=4),
    ], равные=True, растягивать=False)

    группа_фильтров = группа("ГруппаФильтры", 5, [ряд1, ряд2, ряд3],
                             горизонтальная=False, растягивать=False)

    колонки_док = [
        поле("ДокументыПометка", "Объект.Документы.Пометка", 8, тип="CheckBoxField"),
        поле("ДокументыДатаДок", "Объект.Документы.ДатаДок", 8),
        поле("ДокументыДокумент", "Объект.Документы.Документ", 8),
        поле("ДокументыОрганизация", "Объект.Документы.Организация", 8),
        поле("ДокументыСклад", "Объект.Документы.Склад", 8),
        поле("ДокументыПодразделение", "Объект.Документы.Подразделение", 8),
        поле("ДокументыСуммаПоПроводкам", "Объект.Документы.СуммаПоПроводкам", 8,
             только_чтение=True),
    ]
    табл_док = таблица("Документы", "Объект.Документы", 5, колонки_док, титул="Документи")

    стр_документы = страница("СтраницаДокументы", "Документи", 4,
                             [группа_фильтров, табл_док])

    # ---------------- страница «Шапка і комісія» ----------------
    шапка_акта = группа("ГруппаШапкаАкта", 6, [
        поле("НаименованиеРаботы", "Объект.НаименованиеРаботы", 8),
        поле("НомерАкта", "Объект.НомерАкта", 8),
    ], горизонтальная=False, растягивать=False)

    # TitleLocation=Top обязателен: при Left подпись+значение не влезают -> горизонтальная
    # прокрутка формы (боевая грабля документа РасчетКомплектаций, коммит 1bd6252f6a)
    комиссия = группа("ГруппаКомиссия", 6, [
        поле("Бухгалтер", "Объект.Бухгалтер", 8, сверху=True),
        поле("Проверил", "Объект.Проверил", 8, сверху=True),
        поле("НачальникВТО", "Объект.НачальникВТО", 8, сверху=True),
        поле("Инженер", "Объект.Инженер", 8, сверху=True),
        поле("Начальник", "Объект.Начальник", 8, сверху=True),
        поле("МОЛ", "Объект.МОЛ", 8, сверху=True),
    ], горизонтальная=False, титул="Комісія (підписи М-29)", растягивать=False)

    табл_счета = таблица("СчетаМШП", "Объект.СчетаМШП", 6,
                         [поле("СчетаМШПСчет", "Объект.СчетаМШП.Счет", 8)],
                         титул="Рахунки МШП", высота=3)
    группа_счетов = группа("ГруппаСчетаМШП", 5, [табл_счета],
                           горизонтальная=False, растягивать=False)

    стр_шапка = страница("СтраницаШапка", "Шапка і комісія", 4,
                         [шапка_акта, комиссия, группа_счетов])

    # ---------------- страницы ----------------
    ид_страниц = nid()
    ид_подсказки_страниц = nid()
    страницы = (f'\t\t<Pages name="Страницы" id="{ид_страниц}">\n'
                f"\t\t\t<PagesRepresentation>TabsOnTop</PagesRepresentation>\n"
                f'\t\t\t<ExtendedTooltip name="СтраницыРасширеннаяПодсказка" '
                f'id="{ид_подсказки_страниц}"/>\n'
                f"\t\t\t<ChildItems>\n{стр_документы}\n{стр_шапка}\n\t\t\t</ChildItems>\n"
                f"\t\t</Pages>")

    # ---------------- команды ----------------
    команды = []
    for имя, титул in (("Заполнить", "Заповнити"), ("ОтметитьВсе", "Відмітити всі"),
                       ("СнятьОтметки", "Зняти відмітки"), ("ПечатьМ29", "Друк М-29")):
        команды.append(f'\t\t<Command name="{имя}" id="{nid()}">\n'
                       f"{назва(титул, chr(9) * 3)}\n"
                       f"\t\t\t<Action>{имя}</Action>\n\t\t</Command>")

    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<Form {NS_FORM}>
{назва("Друк М-29", chr(9))}
\t<AutoSaveDataInSettings>DontUse</AutoSaveDataInSettings>
\t<AutoTitle>false</AutoTitle>
\t<VerticalScroll>useIfNecessary</VerticalScroll>
\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1">
\t\t<HorizontalAlign>Right</HorizontalAlign>
\t\t<Autofill>false</Autofill>
\t\t<ChildItems>
{chr(10).join(кнопки)}
\t\t</ChildItems>
\t</AutoCommandBar>
\t<Events>
\t\t<Event name="OnCreateAtServer">ПриСозданииНаСервере</Event>
\t\t<Event name="OnClose">ПриЗакрытии</Event>
\t</Events>
\t<ChildItems>
{страницы}
\t</ChildItems>
\t<Attributes>
\t\t<Attribute name="Объект" id="{nid()}">
\t\t\t<Type>
\t\t\t\t<v8:Type>cfg:ExternalDataProcessorObject.{ИМЯ_ОБРАБОТКИ}</v8:Type>
\t\t\t</Type>
\t\t\t<MainAttribute>true</MainAttribute>
\t\t</Attribute>
\t</Attributes>
\t<Commands>
{chr(10).join(команды)}
\t</Commands>
</Form>"""

    with io.open(ФАЙЛ_ФОРМЫ, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write(xml)

    # --- обгортка: uuid СТАБІЛЬНИЙ (переюзаємо наявний, інакше при частковому
    #     завантаженні «нельзя добавлять объекты метаданных без родительского») ---
    uuid_обертки = None
    if os.path.exists(ФАЙЛ_ОБЕРТКИ):
        стар = io.open(ФАЙЛ_ОБЕРТКИ, encoding="utf-8-sig").read()
        m = re.search(r'<Form uuid="([0-9a-f-]+)"', стар)
        if m:
            uuid_обертки = m.group(1)
    if not uuid_обертки:
        uuid_обертки = str(uuid.uuid4())

    обертка = f"""<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject {NS_MD}>
\t<Form uuid="{uuid_обертки}">
\t\t<Properties>
\t\t\t<Name>Форма</Name>
\t\t\t<Synonym>
\t\t\t\t<v8:item>
\t\t\t\t\t<v8:lang>uk</v8:lang>
\t\t\t\t\t<v8:content>Форма</v8:content>
\t\t\t\t</v8:item>
\t\t\t</Synonym>
\t\t\t<Comment/>
\t\t\t<FormType>Managed</FormType>
\t\t\t<IncludeHelpInContents>false</IncludeHelpInContents>
\t\t\t<UsePurposes>
\t\t\t\t<v8:Value xsi:type="app:ApplicationUsePurpose">PlatformApplication</v8:Value>
\t\t\t\t<v8:Value xsi:type="app:ApplicationUsePurpose">MobilePlatformApplication</v8:Value>
\t\t\t</UsePurposes>
\t\t\t<ExtendedPresentation/>
\t\t</Properties>
\t</Form>
</MetaDataObject>"""

    with io.open(ФАЙЛ_ОБЕРТКИ, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write(обертка)

    print(f"Записано: {ФАЙЛ_ФОРМЫ}")
    print(f"Записано: {ФАЙЛ_ОБЕРТКИ}  (uuid обгортки {uuid_обертки})")
    print(f"  максимальний id елемента: {_id[0]}")


if __name__ == "__main__":
    main()
