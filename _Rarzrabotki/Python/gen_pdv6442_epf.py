# -*- coding: utf-8 -*-
"""
Генератор XML-исходників зовнішньої обробки А_РозрахункиПДВЗаПершоюПодією (BAS Бухгалтерія, формат 2.13).

Пише: корневий XML, Forms/Форма.xml, Forms/Форма/Ext/Form.xml, Templates/Инструкция.xml.
Модулі (.bsl) та HTML-інструкція пишуться окремо (руками) і генератор їх НЕ чіпає.
GUID детерміновані (uuid5 від імені об'єкта і шляху) — повторний запуск не змінює ідентифікатори,
перетину з іншими обробками немає. ClassId зовнішньої обробки — константа класу.

Запуск:  python gen_pdv6442_epf.py [каталог_Обработки]
"""
import os, sys, uuid

NAME = 'А_РозрахункиПДВЗаПершоюПодією'
SYN = 'Розрахунки ПДВ за першою подією (вирівнювання 6442)'
CLASS_ID = 'c3831ec8-d8d5-4f93-8a22-f9bfae07327f'
NS = uuid.uuid5(uuid.NAMESPACE_URL, 'industrialbud/epf/' + NAME)


def gid(path):
    return str(uuid.uuid5(NS, path))


XMLNS_MD = ('xmlns="http://v8.1c.ru/8.3/MDClasses" xmlns:app="http://v8.1c.ru/8.2/managed-application/core" '
            'xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" xmlns:cmi="http://v8.1c.ru/8.2/managed-application/cmi" '
            'xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" '
            'xmlns:style="http://v8.1c.ru/8.1/data/ui/style" xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" '
            'xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" '
            'xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" '
            'xmlns:xen="http://v8.1c.ru/8.3/xcf/enums" xmlns:xpr="http://v8.1c.ru/8.3/xcf/predef" '
            'xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" xmlns:xs="http://www.w3.org/2001/XMLSchema" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13"')
XMLNS_FORM = ('xmlns="http://v8.1c.ru/8.3/xcf/logform" xmlns:app="http://v8.1c.ru/8.2/managed-application/core" '
              'xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" '
              'xmlns:dcscor="http://v8.1c.ru/8.1/data-composition-system/core" '
              'xmlns:dcsset="http://v8.1c.ru/8.1/data-composition-system/settings" '
              'xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" '
              'xmlns:style="http://v8.1c.ru/8.1/data/ui/style" xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" '
              'xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" '
              'xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" '
              'xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" xmlns:xs="http://www.w3.org/2001/XMLSchema" '
              'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13"')


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def loc(tag, text, ind):
    """Багатомовний рядок: ru і uk однакові (база українська — порожній uk дає порожній заголовок)."""
    t = '\t' * ind
    return (f'{t}<{tag}>\n'
            f'{t}\t<v8:item>\n{t}\t\t<v8:lang>ru</v8:lang>\n{t}\t\t<v8:content>{esc(text)}</v8:content>\n{t}\t</v8:item>\n'
            f'{t}\t<v8:item>\n{t}\t\t<v8:lang>uk</v8:lang>\n{t}\t\t<v8:content>{esc(text)}</v8:content>\n{t}\t</v8:item>\n'
            f'{t}</{tag}>\n')


# ---------------------------------------------------------------- типи
def T(kind, *a):
    return (kind,) + a


def type_xml(tp, ind):
    t = '\t' * ind
    k = tp[0]
    if k == 'bool':
        body = f'{t}\t<v8:Type>xs:boolean</v8:Type>\n'
    elif k == 'str':
        ln = tp[1]
        body = (f'{t}\t<v8:Type>xs:string</v8:Type>\n{t}\t<v8:StringQualifiers>\n{t}\t\t<v8:Length>{ln}</v8:Length>\n'
                f'{t}\t\t<v8:AllowedLength>Variable</v8:AllowedLength>\n{t}\t</v8:StringQualifiers>\n')
    elif k == 'num':
        d, f = tp[1], tp[2]
        body = (f'{t}\t<v8:Type>xs:decimal</v8:Type>\n{t}\t<v8:NumberQualifiers>\n{t}\t\t<v8:Digits>{d}</v8:Digits>\n'
                f'{t}\t\t<v8:FractionDigits>{f}</v8:FractionDigits>\n{t}\t\t<v8:AllowedSign>Any</v8:AllowedSign>\n'
                f'{t}\t</v8:NumberQualifiers>\n')
    elif k == 'date':
        fr = tp[1] if len(tp) > 1 else 'Date'
        body = (f'{t}\t<v8:Type>xs:dateTime</v8:Type>\n{t}\t<v8:DateQualifiers>\n'
                f'{t}\t\t<v8:DateFractions>{fr}</v8:DateFractions>\n{t}\t</v8:DateQualifiers>\n')
    elif k == 'ref':
        body = f'{t}\t<v8:Type>cfg:{tp[1]}</v8:Type>\n'
    elif k == 'docref':
        body = f'{t}\t<v8:TypeSet>cfg:DocumentRef</v8:TypeSet>\n'
    else:
        raise ValueError(k)
    return f'{t}<Type>\n{body}{t}</Type>\n'


# ---------------------------------------------------------------- метадані
ATTRS = [
    ('Организация', 'Організація (порожньо — усі)', T('ref', 'CatalogRef.Организации')),
    ('ДатаНачала', 'Період з', T('date', 'Date')),
    ('ДатаОкончания', 'Період по', T('date', 'Date')),
    ('Контрагент', 'Контрагент', T('ref', 'CatalogRef.Контрагенты')),
    ('ДоговорКонтрагента', 'Договір', T('ref', 'CatalogRef.ДоговорыКонтрагентов')),
    ('ГраницаЗапрета', 'Дата заборони змін (для всіх)', T('date', 'Date')),
    ('РазрешитьЗакрытыйПериод', 'Дозволити закритий період', T('bool')),
    ('ТолькоРасхождения', 'Лише документи з розбіжністю', T('bool')),
]

TS = {
    'Ключи': ('Ключі розрахунків', [
        ('Отметка', 'V', T('bool')),
        ('КлючИД', 'Ключ', T('str', 300)),
        ('Организация', 'Організація', T('ref', 'CatalogRef.Организации')),
        ('Контрагент', 'Контрагент', T('ref', 'CatalogRef.Контрагенты')),
        ('ДоговорКонтрагента', 'Договір', T('ref', 'CatalogRef.ДоговорыКонтрагентов')),
        ('Сделка', 'Угода / рахунок (замовлення)', T('docref')),
        ('СтавкаНДС', 'Ставка ПДВ', T('ref', 'EnumRef.СтавкиНДС')),
        ('ВозвратнаяТара', 'Зворотна тара', T('bool')),
        ('ДляХозяйственнойДеятельности', 'Для госп. діяльності', T('bool')),
        ('ВидДеятельностиНДС', 'Вид діяльності ПДВ', T('ref', 'EnumRef.ВидыДеятельностиНДС')),
        ('Амортизируется', 'Амортизується', T('bool')),
        ('ОжидаемоПС', 'Очікувано ПС', T('num', 15, 2)),
        ('ФактПС', 'Факт ПС', T('num', 15, 2)),
        ('Отклонение', 'Зайве (факт − очікувано)', T('num', 15, 2)),
        ('Сальдо6442', 'Сальдо 6442', T('num', 15, 2)),
        ('ДокументовВсего', 'Документів', T('num', 6, 0)),
        ('ДокументовСРасхождением', 'З розбіжністю', T('num', 6, 0)),
        ('ПервыйДокумент', 'Перший документ з розбіжністю', T('docref')),
        ('ЗакрытыйПериод', 'Закритий період', T('bool')),
        ('Статус', 'Статус', T('str', 0)),
    ]),
    'Документы': ('Документи ключа', [
        ('Отметка', 'V', T('bool')),
        ('КлючИД', 'Ключ', T('str', 300)),
        ('Документ', 'Документ', T('docref')),
        ('Дата', 'Дата', T('date', 'DateTime')),
        ('Событие', 'Подія', T('str', 50)),
        ('НДСПоступление', 'ПДВ надходження', T('num', 15, 2)),
        ('НДСОплата', 'ПДВ оплати', T('num', 15, 2)),
        ('ОжидаемоПС', 'Очікувано ПС', T('num', 15, 2)),
        ('ФактПС', 'Факт ПС', T('num', 15, 2)),
        ('Отклонение', 'Різниця', T('num', 15, 2)),
        ('ЕстьРасхождение', 'Є розбіжність', T('bool')),
        ('ЗакрытыйПериод', 'Закритий період', T('bool')),
        ('Статус', 'Статус', T('str', 0)),
    ]),
    'Протокол': ('Протокол перепроведення', [
        ('Документ', 'Документ', T('docref')),
        ('Дата', 'Дата', T('date', 'DateTime')),
        ('ВидДокумента', 'Вид документа', T('str', 100)),
        ('Статус', 'Статус', T('str', 0)),
        ('БылВОчередиОбмена', 'Був у черзі обміну', T('bool')),
        ('СнятСОчереди', 'Знято з черги', T('bool')),
    ]),
}


def attr_xml(path, name, syn, tp, ind, in_ts):
    t = '\t' * ind
    s = f'{t}<Attribute uuid="{gid(path)}">\n{t}\t<Properties>\n{t}\t\t<Name>{name}</Name>\n'
    s += loc('Synonym', syn, ind + 2)
    s += f'{t}\t\t<Comment/>\n'
    s += type_xml(tp, ind + 2)
    s += (f'{t}\t\t<PasswordMode>false</PasswordMode>\n{t}\t\t<Format/>\n{t}\t\t<EditFormat/>\n{t}\t\t<ToolTip/>\n'
          f'{t}\t\t<MarkNegatives>{"true" if tp[0] == "num" else "false"}</MarkNegatives>\n{t}\t\t<Mask/>\n'
          f'{t}\t\t<MultiLine>false</MultiLine>\n{t}\t\t<ExtendedEdit>false</ExtendedEdit>\n'
          f'{t}\t\t<MinValue xsi:nil="true"/>\n{t}\t\t<MaxValue xsi:nil="true"/>\n')
    if in_ts:
        s += f'{t}\t\t<FillFromFillingValue>false</FillFromFillingValue>\n{t}\t\t<FillValue xsi:nil="true"/>\n'
    s += (f'{t}\t\t<FillChecking>DontCheck</FillChecking>\n{t}\t\t<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>\n'
          f'{t}\t\t<ChoiceParameterLinks/>\n{t}\t\t<ChoiceParameters/>\n{t}\t\t<QuickChoice>Auto</QuickChoice>\n'
          f'{t}\t\t<CreateOnInput>Auto</CreateOnInput>\n{t}\t\t<ChoiceForm/>\n{t}\t\t<LinkByType/>\n'
          f'{t}\t\t<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>\n{t}\t</Properties>\n{t}</Attribute>\n')
    return s


def root_xml():
    s = f'﻿<?xml version="1.0" encoding="UTF-8"?>\n<MetaDataObject {XMLNS_MD}>\n'
    s += f'\t<ExternalDataProcessor uuid="{gid("root")}">\n\t\t<InternalInfo>\n\t\t\t<xr:ContainedObject>\n'
    s += f'\t\t\t\t<xr:ClassId>{CLASS_ID}</xr:ClassId>\n\t\t\t\t<xr:ObjectId>{gid("object")}</xr:ObjectId>\n\t\t\t</xr:ContainedObject>\n'
    s += (f'\t\t\t<xr:GeneratedType name="ExternalDataProcessorObject.{NAME}" category="Object">\n'
          f'\t\t\t\t<xr:TypeId>{gid("type/Object")}</xr:TypeId>\n\t\t\t\t<xr:ValueId>{gid("value/Object")}</xr:ValueId>\n'
          f'\t\t\t</xr:GeneratedType>\n\t\t</InternalInfo>\n\t\t<Properties>\n\t\t\t<Name>{NAME}</Name>\n')
    s += loc('Synonym', SYN, 3)
    s += (f'\t\t\t<Comment/>\n\t\t\t<DefaultForm>ExternalDataProcessor.{NAME}.Form.Форма</DefaultForm>\n'
          f'\t\t\t<AuxiliaryForm/>\n\t\t</Properties>\n\t\t<ChildObjects>\n')
    for n, syn, tp in ATTRS:
        s += attr_xml('attr/' + n, n, syn, tp, 3, False)
    for ts_name, (ts_syn, cols) in TS.items():
        s += f'\t\t\t<TabularSection uuid="{gid("ts/" + ts_name)}">\n\t\t\t\t<InternalInfo>\n'
        for cat, pref in (('TabularSection', 'ExternalDataProcessorTabularSection'),
                          ('TabularSectionRow', 'ExternalDataProcessorTabularSectionRow')):
            s += (f'\t\t\t\t\t<xr:GeneratedType name="{pref}.{NAME}.{ts_name}" category="{cat}">\n'
                  f'\t\t\t\t\t\t<xr:TypeId>{gid("type/" + cat + "/" + ts_name)}</xr:TypeId>\n'
                  f'\t\t\t\t\t\t<xr:ValueId>{gid("value/" + cat + "/" + ts_name)}</xr:ValueId>\n'
                  f'\t\t\t\t\t</xr:GeneratedType>\n')
        s += f'\t\t\t\t</InternalInfo>\n\t\t\t\t<Properties>\n\t\t\t\t\t<Name>{ts_name}</Name>\n'
        s += loc('Synonym', ts_syn, 5)
        s += '\t\t\t\t\t<Comment/>\n\t\t\t\t\t<ToolTip/>\n\t\t\t\t\t<FillChecking>DontCheck</FillChecking>\n\t\t\t\t</Properties>\n\t\t\t\t<ChildObjects>\n'
        for n, syn, tp in cols:
            s += attr_xml(f'ts/{ts_name}/{n}', n, syn, tp, 5, True)
        s += '\t\t\t\t</ChildObjects>\n\t\t\t</TabularSection>\n'
    s += '\t\t\t<Form>Форма</Form>\n\t\t\t<Template>Инструкция</Template>\n\t\t</ChildObjects>\n\t</ExternalDataProcessor>\n</MetaDataObject>\n'
    return s


def form_md_xml():
    s = f'﻿<?xml version="1.0" encoding="UTF-8"?>\n<MetaDataObject {XMLNS_MD}>\n\t<Form uuid="{gid("form/Форма")}">\n'
    s += '\t\t<Properties>\n\t\t\t<Name>Форма</Name>\n'
    s += loc('Synonym', 'Форма', 3)
    s += ('\t\t\t<Comment/>\n\t\t\t<FormType>Managed</FormType>\n\t\t\t<IncludeHelpInContents>false</IncludeHelpInContents>\n'
          '\t\t\t<UsePurposes>\n\t\t\t\t<v8:Value xsi:type="app:ApplicationUsePurpose">PlatformApplication</v8:Value>\n'
          '\t\t\t\t<v8:Value xsi:type="app:ApplicationUsePurpose">MobilePlatformApplication</v8:Value>\n'
          '\t\t\t</UsePurposes>\n\t\t\t<ExtendedPresentation/>\n\t\t</Properties>\n\t</Form>\n</MetaDataObject>\n')
    return s


def template_md_xml():
    s = f'﻿<?xml version="1.0" encoding="UTF-8"?>\n<MetaDataObject {XMLNS_MD}>\n\t<Template uuid="{gid("template/Инструкция")}">\n'
    s += '\t\t<Properties>\n\t\t\t<Name>Инструкция</Name>\n'
    s += loc('Synonym', 'Інструкція (HTML)', 3)
    s += '\t\t\t<Comment/>\n\t\t\t<TemplateType>TextDocument</TemplateType>\n\t\t</Properties>\n\t</Template>\n</MetaDataObject>\n'
    return s


# ---------------------------------------------------------------- форма
class Ids:
    def __init__(self):
        self.n = 0

    def __call__(self):
        self.n += 1
        return self.n


ID = Ids()


def cm_et(name, ind):
    t = '\t' * ind
    return (f'{t}<ContextMenu name="{name}КонтекстноеМеню" id="{ID()}"/>\n'
            f'{t}<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ID()}"/>\n')


def events_xml(events, ind):
    if not events:
        return ''
    t = '\t' * ind
    s = f'{t}<Events>\n'
    for ev, h in events:
        s += f'{t}\t<Event name="{ev}">{h}</Event>\n'
    return s + f'{t}</Events>\n'


def input_field(name, path, title=None, ind=0, ro=False, width=None, footer=None, events=None, title_loc=None,
                choice_button=None, hstretch=None):
    t = '\t' * ind
    s = f'{t}<InputField name="{name}" id="{ID()}">\n{t}\t<DataPath>{path}</DataPath>\n'
    if ro:
        s += f'{t}\t<ReadOnly>true</ReadOnly>\n'
    if title:
        s += loc('Title', title, ind + 1)
    if title_loc:
        s += f'{t}\t<TitleLocation>{title_loc}</TitleLocation>\n'
    if footer:
        s += f'{t}\t<FooterDataPath>{footer}</FooterDataPath>\n'
    if width:
        s += f'{t}\t<Width>{width}</Width>\n'
    if hstretch is not None:
        s += f'{t}\t<HorizontalStretch>{hstretch}</HorizontalStretch>\n'
    if choice_button is not None:
        s += f'{t}\t<ChoiceButton>{choice_button}</ChoiceButton>\n'
    s += cm_et(name, ind + 1)
    s += events_xml(events, ind + 1)
    return s + f'{t}</InputField>\n'


def check_field(name, path, title, ind, events=None, ro=False, title_loc=None):
    t = '\t' * ind
    s = f'{t}<CheckBoxField name="{name}" id="{ID()}">\n{t}\t<DataPath>{path}</DataPath>\n'
    if ro:
        s += f'{t}\t<ReadOnly>true</ReadOnly>\n'
    s += loc('Title', title, ind + 1)
    if title_loc:
        s += f'{t}\t<TitleLocation>{title_loc}</TitleLocation>\n'
    s += f'{t}\t<CheckBoxType>Auto</CheckBoxType>\n'
    s += cm_et(name, ind + 1)
    s += events_xml(events, ind + 1)
    return s + f'{t}</CheckBoxField>\n'


def button(name, cmd, ind, kind='CommandBarButton', rep=None, default=False, picture=None, title=None):
    t = '\t' * ind
    s = f'{t}<Button name="{name}" id="{ID()}">\n{t}\t<Type>{kind}</Type>\n'
    if rep:
        s += f'{t}\t<Representation>{rep}</Representation>\n'
    if default:
        s += f'{t}\t<DefaultButton>true</DefaultButton>\n'
    s += f'{t}\t<CommandName>Form.Command.{cmd}</CommandName>\n'
    if title:
        s += loc('Title', title, ind + 1)
    if picture:
        s += f'{t}\t<Picture>\n{t}\t\t<xr:Ref>{picture}</xr:Ref>\n{t}\t\t<xr:LoadTransparent>true</xr:LoadTransparent>\n{t}\t</Picture>\n'
    s += f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ID()}"/>\n'
    return s + f'{t}</Button>\n'


def button_group(name, title, buttons, ind):
    t = '\t' * ind
    s = f'{t}<ButtonGroup name="{name}" id="{ID()}">\n'
    s += loc('Title', title, ind + 1)
    s += f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ID()}"/>\n{t}\t<ChildItems>\n'
    for b in buttons:
        s += button(*b, ind=ind + 2) if isinstance(b, tuple) and len(b) == 2 else button(b[0], b[1], ind + 2, picture=b[2])
    return s + f'{t}\t</ChildItems>\n{t}</ButtonGroup>\n'


def usual_group(name, title, children, ind, group='AlwaysHorizontal', show_title=False, extra=''):
    t = '\t' * ind
    s = f'{t}<UsualGroup name="{name}" id="{ID()}">\n'
    s += loc('Title', title, ind + 1)
    s += f'{t}\t<Group>{group}</Group>\n{t}\t<Behavior>Usual</Behavior>\n'
    s += f'{t}\t<ShowTitle>{"true" if show_title else "false"}</ShowTitle>\n' + extra
    s += f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ID()}"/>\n{t}\t<ChildItems>\n'
    s += ''.join(children)
    return s + f'{t}\t</ChildItems>\n{t}</UsualGroup>\n'


def table(name, path, title, bar_children, columns, events, ind, footer=False, height=None, title_loc=None):
    t = '\t' * ind
    s = f'{t}<Table name="{name}" id="{ID()}">\n'
    s += f'{t}\t<ChangeRowSet>false</ChangeRowSet>\n{t}\t<ChangeRowOrder>false</ChangeRowOrder>\n'
    if height:
        s += f'{t}\t<Height>{height}</Height>\n'
    if footer:
        s += f'{t}\t<Footer>true</Footer>\n'
    s += f'{t}\t<DataPath>{path}</DataPath>\n'
    s += loc('Title', title, ind + 1)
    if title_loc:
        s += f'{t}\t<TitleLocation>{title_loc}</TitleLocation>\n'
    s += f'{t}\t<RowFilter xsi:nil="true"/>\n'
    s += f'{t}\t<ContextMenu name="{name}КонтекстноеМеню" id="{ID()}"/>\n'
    s += f'{t}\t<AutoCommandBar name="{name}КоманднаяПанель" id="{ID()}">\n{t}\t\t<Autofill>false</Autofill>\n'
    if bar_children:
        s += f'{t}\t\t<ChildItems>\n' + ''.join(bar_children) + f'{t}\t\t</ChildItems>\n'
    s += f'{t}\t</AutoCommandBar>\n'
    s += f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ID()}"/>\n'
    for add, typ, suf in (('SearchStringAddition', 'SearchStringRepresentation', 'СтрокаПоиска'),
                          ('ViewStatusAddition', 'ViewStatusRepresentation', 'СостояниеПросмотра'),
                          ('SearchControlAddition', 'SearchControl', 'УправлениеПоиском')):
        s += (f'{t}\t<{add} name="{name}{suf}" id="{ID()}">\n{t}\t\t<AdditionSource>\n{t}\t\t\t<Item>{name}</Item>\n'
              f'{t}\t\t\t<Type>{typ}</Type>\n{t}\t\t</AdditionSource>\n'
              f'{t}\t\t<ContextMenu name="{name}{suf}КонтекстноеМеню" id="{ID()}"/>\n'
              f'{t}\t\t<ExtendedTooltip name="{name}{suf}РасширеннаяПодсказка" id="{ID()}"/>\n{t}\t</{add}>\n')
    s += events_xml(events, ind + 1)
    s += f'{t}\t<ChildItems>\n' + ''.join(columns) + f'{t}\t</ChildItems>\n'
    return s + f'{t}</Table>\n'


def page(name, title, children, ind, picture=None):
    t = '\t' * ind
    s = f'{t}<Page name="{name}" id="{ID()}">\n'
    s += loc('Title', title, ind + 1)
    if picture:
        s += f'{t}\t<Picture>\n{t}\t\t<xr:Ref>{picture}</xr:Ref>\n{t}\t\t<xr:LoadTransparent>true</xr:LoadTransparent>\n{t}\t</Picture>\n'
    s += f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ID()}"/>\n{t}\t<ChildItems>\n'
    s += ''.join(children)
    return s + f'{t}\t</ChildItems>\n{t}</Page>\n'


def label_field(name, path, ind, hstretch=True):
    t = '\t' * ind
    s = (f'{t}<LabelField name="{name}" id="{ID()}">\n{t}\t<DataPath>{path}</DataPath>\n'
         f'{t}\t<TitleLocation>None</TitleLocation>\n{t}\t<AutoMaxWidth>false</AutoMaxWidth>\n'
         f'{t}\t<HorizontalStretch>{"true" if hstretch else "false"}</HorizontalStretch>\n')
    s += cm_et(name, ind + 1)
    return s + f'{t}</LabelField>\n'


def form_xml():
    ID.n = 0
    body = []
    # --- шапка
    i = 3
    hdr1 = [
        input_field('Организация', 'Объект.Организация', 'Організація', i + 1, width=28),
        input_field('ДатаНачала', 'Объект.ДатаНачала', 'Період з', i + 1, width=10),
        input_field('ДатаОкончания', 'Объект.ДатаОкончания', 'по', i + 1, width=10),
        button('ВыбратьПериодКнопка', 'ВыбратьПериод', i + 1, kind='UsualButton', title='...'),
        input_field('ГраницаЗапрета', 'Объект.ГраницаЗапрета', 'Дата заборони (для всіх)', i + 1, ro=True, width=10),
    ]
    hdr2 = [
        input_field('Контрагент', 'Объект.Контрагент', 'Контрагент', i + 1, width=28,
                    events=[('OnChange', 'КонтрагентПриИзменении')]),
        input_field('ДоговорКонтрагента', 'Объект.ДоговорКонтрагента', 'Договір', i + 1, width=28),
        check_field('РазрешитьЗакрытыйПериод', 'Объект.РазрешитьЗакрытыйПериод', 'Дозволити закритий період', i + 1,
                    events=[('OnChange', 'РазрешитьЗакрытыйПериодПриИзменении')]),
        check_field('ТолькоРасхождения', 'Объект.ТолькоРасхождения', 'Лише розбіжності', i + 1,
                    events=[('OnChange', 'ТолькоРасхожденияПриИзменении')]),
    ]
    body.append(usual_group('ГруппаШапка', 'Параметри', [
        usual_group('ГруппаШапка1', 'Рядок 1', hdr1, i + 2),
        usual_group('ГруппаШапка2', 'Рядок 2', hdr2, i + 2),
    ], 2, group='Vertical'))
    body.append(label_field('НадписьРезюме', 'НадписьРезюме', 2))

    # --- сторінка «Проблеми»: ключі
    p = 5
    kl_bar = [
        button('КлючиПерепровестиОтмеченные', 'ПерепровестиОтмеченные', p + 2, rep='PictureAndText',
               picture='StdPicture.ExecuteTask'),
        button_group('КлючиГруппаФлаги', 'Прапори', [
            ('КлючиОтметитьВсе', 'ОтметитьВсеКлючи', 'StdPicture.CheckAll'),
            ('КлючиСнятьВсе', 'СнятьВсеКлючи', 'StdPicture.UncheckAll'),
            ('КлючиИнвертировать', 'ИнвертироватьКлючи', None),
        ], p + 2),
        button('КлючиКарточкаСчета', 'КарточкаСчета6442', p + 2, rep='Text'),
    ]
    TSK = 'Объект.Ключи'
    kl_cols = [
        check_field('КлючиОтметка', TSK + '.Отметка', 'V', p + 2, events=[('OnChange', 'КлючиОтметкаПриИзменении')]),
        input_field('КлючиОрганизация', TSK + '.Организация', 'Організація', p + 2, ro=True, width=12),
        input_field('КлючиКонтрагент', TSK + '.Контрагент', 'Контрагент', p + 2, ro=True, width=20),
        input_field('КлючиДоговор', TSK + '.ДоговорКонтрагента', 'Договір', p + 2, ro=True, width=20),
        input_field('КлючиСделка', TSK + '.Сделка', 'Угода / рахунок', p + 2, ro=True, width=24),
        input_field('КлючиСтавка', TSK + '.СтавкаНДС', 'Ставка', p + 2, ro=True, width=5),
        input_field('КлючиОжидаемо', TSK + '.ОжидаемоПС', 'Очікувано ПС', p + 2, ro=True, width=11,
                    footer=TSK + '.TotalОжидаемоПС'),
        input_field('КлючиФакт', TSK + '.ФактПС', 'Факт ПС', p + 2, ro=True, width=11, footer=TSK + '.TotalФактПС'),
        input_field('КлючиОтклонение', TSK + '.Отклонение', 'Зайве', p + 2, ro=True, width=11,
                    footer=TSK + '.TotalОтклонение'),
        input_field('КлючиСальдо6442', TSK + '.Сальдо6442', 'Сальдо 6442', p + 2, ro=True, width=11,
                    footer=TSK + '.TotalСальдо6442'),
        input_field('КлючиДокСРасхождением', TSK + '.ДокументовСРасхождением', 'Розб.', p + 2, ro=True, width=5,
                    footer=TSK + '.TotalДокументовСРасхождением'),
        input_field('КлючиДокВсего', TSK + '.ДокументовВсего', 'Док.', p + 2, ro=True, width=5),
        input_field('КлючиПервыйДокумент', TSK + '.ПервыйДокумент', 'Перший документ з розбіжністю', p + 2, ro=True,
                    width=24),
        check_field('КлючиЗакрытыйПериод', TSK + '.ЗакрытыйПериод', 'Закр.', p + 2, ro=True),
        input_field('КлючиСтатус', TSK + '.Статус', 'Статус', p + 2, ro=True, width=20),
    ]
    t_keys = table('Ключи', TSK, 'Ключі з розбіжністю (договір × угода × ставка)', kl_bar, kl_cols,
                   [('Selection', 'КлючиВыбор'), ('OnActivateRow', 'КлючиПриАктивизацииСтроки')], p, footer=True,
                   height=8)
    TSD = 'Объект.Документы'
    dk_bar = [
        button_group('ДокументыГруппаФлаги', 'Прапори', [
            ('ДокументыОтметитьВыделенные', 'ОтметитьВыделенныеДокументы', 'StdPicture.CheckAll'),
            ('ДокументыСнятьВыделенные', 'СнятьВыделенныеДокументы', 'StdPicture.UncheckAll'),
        ], p + 2),
        button('ДокументыОткрытьДокумент', 'ОткрытьДокумент', p + 2, rep='Text'),
    ]
    dk_cols = [
        check_field('ДокументыОтметка', TSD + '.Отметка', 'V', p + 2,
                    events=[('OnChange', 'ДокументыОтметкаПриИзменении')]),
        input_field('ДокументыДокумент', TSD + '.Документ', 'Документ', p + 2, ro=True, width=34),
        input_field('ДокументыСобытие', TSD + '.Событие', 'Подія', p + 2, ro=True, width=12),
        input_field('ДокументыНДСПоступление', TSD + '.НДСПоступление', 'ПДВ надходж.', p + 2, ro=True, width=11),
        input_field('ДокументыНДСОплата', TSD + '.НДСОплата', 'ПДВ оплати', p + 2, ro=True, width=11),
        input_field('ДокументыОжидаемо', TSD + '.ОжидаемоПС', 'Очікувано ПС', p + 2, ro=True, width=11),
        input_field('ДокументыФакт', TSD + '.ФактПС', 'Факт ПС', p + 2, ro=True, width=11),
        input_field('ДокументыОтклонение', TSD + '.Отклонение', 'Різниця', p + 2, ro=True, width=11),
        check_field('ДокументыЗакрытыйПериод', TSD + '.ЗакрытыйПериод', 'Закр.', p + 2, ro=True),
        input_field('ДокументыСтатус', TSD + '.Статус', 'Статус', p + 2, ro=True, width=24),
    ]
    t_docs = table('Документы', TSD, 'Документи ключа (хронологічно)', dk_bar, dk_cols, [('Selection', 'ДокументыВыбор')],
                   p, height=7)
    pg1 = page('СтраницаПроблемы', 'Проблеми', [t_keys, t_docs], 4, picture='StdPicture.Report')

    # --- сторінка «Перепроведення за період»
    TSP = 'Объект.Протокол'
    per_bar = [
        button('ПротоколПоказатьДокументыПериода', 'ПоказатьДокументыПериода', p + 2, rep='Text'),
        button('ПротоколПерепровестиПериод', 'ПерепровестиПериод', p + 2, rep='PictureAndText',
               picture='StdPicture.ExecuteTask'),
    ]
    pr_cols = [
        input_field('ПротоколДокумент', TSP + '.Документ', 'Документ', p + 2, ro=True, width=40),
        input_field('ПротоколВидДокумента', TSP + '.ВидДокумента', 'Вид', p + 2, ro=True, width=18),
        input_field('ПротоколСтатус', TSP + '.Статус', 'Статус', p + 2, ro=True, width=30),
        check_field('ПротоколБылВОчереди', TSP + '.БылВОчередиОбмена', 'Був у черзі', p + 2, ro=True),
        check_field('ПротоколСнятСОчереди', TSP + '.СнятСОчереди', 'Знято з черги', p + 2, ro=True),
    ]
    t_prot = table('Протокол', TSP, 'Документи механізму ПДВ за період / протокол', per_bar, pr_cols,
                   [('Selection', 'ПротоколВыбор')], p)
    pg2 = page('СтраницаПериод', 'Перепроведення за період', [
        label_field('НадписьПериод', 'НадписьПериод', 5), t_prot], 4, picture='StdPicture.DataCompositionSettingsWizard')

    # --- сторінка «Журнал»
    jr_bar = [button('ЖурналОбновить', 'ОбновитьЖурнал', p + 2, rep='PictureAndText', picture='StdPicture.Refresh')]
    jr_cols = [
        input_field('ЖурналДата', 'Журнал.Дата', 'Дата', p + 2, ro=True, width=16),
        input_field('ЖурналУровень', 'Журнал.Уровень', 'Рівень', p + 2, ro=True, width=10),
        input_field('ЖурналКомментарий', 'Журнал.Комментарий', 'Коментар', p + 2, ro=True, width=80),
    ]
    t_jr = table('Журнал', 'Журнал', 'Журнал нічних запусків (30 днів)', jr_bar, jr_cols,
                 [('Selection', 'ЖурналВыбор')], p)
    pg3 = page('СтраницаЖурнал', 'Журнал нічних запусків', [t_jr], 4, picture='StdPicture.EventLog')

    # --- сторінка «Інструкція»
    t = '\t' * 5
    html = (f'{t}<HTMLDocumentField name="ПолеИнструкции" id="{ID()}">\n{t}\t<DataPath>ТекстИнструкции</DataPath>\n'
            f'{t}\t<TitleLocation>None</TitleLocation>\n'
            f'{t}\t<ContextMenu name="ПолеИнструкцииКонтекстноеМеню" id="{ID()}"/>\n'
            f'{t}\t<ExtendedTooltip name="ПолеИнструкцииРасширеннаяПодсказка" id="{ID()}"/>\n{t}</HTMLDocumentField>\n')
    pg4 = page('СтраницаИнструкция', 'Інструкція', [html], 4, picture='StdPicture.Information')

    pages_id = ID()
    pages = (f'\t\t<Pages name="Страницы" id="{pages_id}">\n\t\t\t<HorizontalStretch>true</HorizontalStretch>\n'
             f'\t\t\t<VerticalStretch>true</VerticalStretch>\n\t\t\t<PagesRepresentation>TabsOnTop</PagesRepresentation>\n'
             f'\t\t\t<ExtendedTooltip name="СтраницыРасширеннаяПодсказка" id="{ID()}"/>\n\t\t\t<ChildItems>\n'
             + pg1 + pg2 + pg3 + pg4 + '\t\t\t</ChildItems>\n\t\t</Pages>\n')
    body.append(pages)

    # --- командна панель форми
    bar = (button('ФормаАнализировать', 'Анализировать', 3, rep='PictureAndText', default=True,
                  picture='StdPicture.GenerateReport')
           + button('ФормаПерепровестиОтмеченные', 'ПерепровестиОтмеченные', 3, rep='PictureAndText',
                    picture='StdPicture.ExecuteTask'))

    # --- реквізити форми
    attrs = ('\t<Attributes>\n\t\t<Attribute name="Объект" id="1">\n\t\t\t<Type>\n'
             f'\t\t\t\t<v8:Type>cfg:ExternalDataProcessorObject.{NAME}</v8:Type>\n\t\t\t</Type>\n'
             '\t\t\t<MainAttribute>true</MainAttribute>\n\t\t\t<Save>\n\t\t\t\t<Field>Объект.Организация</Field>\n'
             '\t\t\t</Save>\n\t\t</Attribute>\n')
    aid = 2

    def f_attr(name, title, tp_xml):
        nonlocal aid
        s = f'\t\t<Attribute name="{name}" id="{aid}">\n' + loc('Title', title, 3) + tp_xml + '\t\t</Attribute>\n'
        aid += 1
        return s

    str0 = type_xml(('str', 0), 3)
    attrs += f_attr('НадписьРезюме', 'Резюме', str0)
    attrs += f_attr('НадписьПериод', 'Період', str0)
    attrs += f_attr('ТекстИнструкции', 'Інструкція', str0)
    attrs += f_attr('ДополнительнаяОбработкаСсылка', 'Додаткова обробка',
                    type_xml(('ref', 'CatalogRef.ДополнительныеОтчетыИОбработки'), 3))
    # Журнал — таблиця значень
    s = f'\t\t<Attribute name="Журнал" id="{aid}">\n' + loc('Title', 'Журнал', 3)
    aid += 1
    s += '\t\t\t<Type>\n\t\t\t\t<v8:Type>v8:ValueTable</v8:Type>\n\t\t\t</Type>\n\t\t\t<Columns>\n'
    for cn, ct in (('Дата', ('date', 'DateTime')), ('Уровень', ('str', 30)), ('Комментарий', ('str', 0))):
        s += f'\t\t\t\t<Column name="{cn}" id="{aid}">\n' + loc('Title', cn, 5) + type_xml(ct, 5) + '\t\t\t\t</Column>\n'
        aid += 1
    s += '\t\t\t</Columns>\n\t\t</Attribute>\n'
    attrs += s + '\t</Attributes>\n'

    # --- команди
    CMDS = [
        ('Анализировать', 'Аналізувати', 'Знайти ключі та документи з задвоєнням ПДВ першої події (6442)'),
        ('ПерепровестиОтмеченные', 'Перепровести відмічені', 'Перепровести відмічені документи в хронологічному порядку'),
        ('ОтметитьВсеКлючи', 'Відмітити всі', 'Відмітити всі видимі ключі'),
        ('СнятьВсеКлючи', 'Зняти всі', 'Зняти відмітки з усіх видимих ключів'),
        ('ИнвертироватьКлючи', 'Інвертувати', 'Інвертувати відмітки видимих ключів'),
        ('ОтметитьВыделенныеДокументы', 'Відмітити виділені', 'Відмітити виділені документи'),
        ('СнятьВыделенныеДокументы', 'Зняти виділені', 'Зняти відмітки з виділених документів'),
        ('ОткрытьДокумент', 'Відкрити документ', 'Відкрити поточний документ'),
        ('КарточкаСчета6442', 'Картка 6442 по ключу', 'Відкрити картку рахунку 6442 з відбором по ключу'),
        ('ПоказатьДокументыПериода', 'Показати документи періоду', 'Документи механізму ПДВ за період (без перепроведення)'),
        ('ПерепровестиПериод', 'Перепровести всі документи ПДВ за період', 'Перепровести всі документи механізму ПДВ за період'),
        ('ОбновитьЖурнал', 'Оновити', 'Прочитати журнал реєстрації за 30 днів'),
        ('ВыбратьПериод', '...', 'Вибрати період'),
    ]
    cmds = '\t<Commands>\n'
    for n, (cn, title, tip) in enumerate(CMDS, 1):
        cmds += f'\t\t<Command name="{cn}" id="{n}">\n' + loc('Title', title, 3) + loc('ToolTip', tip, 3)
        cmds += f'\t\t\t<Action>{cn}</Action>\n\t\t</Command>\n'
    cmds += '\t</Commands>\n'

    x = f'﻿<?xml version="1.0" encoding="UTF-8"?>\n<Form {XMLNS_FORM}>\n'
    x += '\t<AutoSaveDataInSettings>Use</AutoSaveDataInSettings>\n'
    x += f'\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1">\n\t\t<Autofill>false</Autofill>\n\t\t<ChildItems>\n{bar}\t\t</ChildItems>\n\t</AutoCommandBar>\n'
    x += events_xml([('OnCreateAtServer', 'ПриСозданииНаСервере'), ('OnOpen', 'ПриОткрытии')], 1)
    x += '\t<ChildItems>\n' + ''.join(body) + '\t</ChildItems>\n'
    x += attrs + cmds + '</Form>\n'
    return x


def write(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8', newline='\r\n') as f:
        f.write(text)


def main():
    base = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), '..', 'Обработки')
    base = os.path.abspath(base)
    write(os.path.join(base, NAME + '.xml'), root_xml())
    write(os.path.join(base, NAME, 'Forms', 'Форма.xml'), form_md_xml())
    write(os.path.join(base, NAME, 'Forms', 'Форма', 'Ext', 'Form.xml'), form_xml())
    write(os.path.join(base, NAME, 'Templates', 'Инструкция.xml'), template_md_xml())
    print('OK:', base)


if __name__ == '__main__':
    main()
