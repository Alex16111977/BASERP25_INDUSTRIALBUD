# -*- coding: utf-8 -*-
"""Генерация Form.xml трёх форм документа А_ПотребаЗавершенняПроекту (BOM + LF, как выгрузка платформы).

ВНИМАНИЕ: перезаписывает формы целиком (в т.ч. поздние ручные правки и правки, внесённые в Конфигураторе).
    python gen_forms.py <каталог конфигурации, напр. worktree>
"""
from pathlib import Path
from xml.sax.saxutils import escape

DOC = "А_ПотребаЗавершенняПроекту"
# Корень выгрузки конфигурации — аргумент: python gen_forms.py <каталог конфигурации>
import sys
ROOT = Path(sys.argv[1]) / "Documents" / DOC / "Forms"

HEAD = ('<?xml version="1.0" encoding="UTF-8"?>\n'
        '<Form xmlns="http://v8.1c.ru/8.3/xcf/logform" xmlns:app="http://v8.1c.ru/8.2/managed-application/core" '
        'xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" xmlns:dcscor="http://v8.1c.ru/8.1/data-composition-system/core" '
        'xmlns:dcsset="http://v8.1c.ru/8.1/data-composition-system/settings" xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" '
        'xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" xmlns:style="http://v8.1c.ru/8.1/data/ui/style" '
        'xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" xmlns:v8="http://v8.1c.ru/8.1/data/core" '
        'xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" '
        'xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" '
        'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13">\n')


class Ids:
    def __init__(self, start=1):
        self.n = start - 1

    def __call__(self):
        self.n += 1
        return self.n


def ml(tag, text, ind):
    t = escape(text)
    return (f"{ind}<{tag}>\n"
            f"{ind}\t<v8:item>\n{ind}\t\t<v8:lang>ru</v8:lang>\n{ind}\t\t<v8:content>{t}</v8:content>\n{ind}\t</v8:item>\n"
            f"{ind}\t<v8:item>\n{ind}\t\t<v8:lang>uk</v8:lang>\n{ind}\t\t<v8:content>{t}</v8:content>\n{ind}\t</v8:item>\n"
            f"{ind}</{tag}>\n")


def input_field(ids, name, path, ind, title=None, extra=""):
    s = f'{ind}<InputField name="{name}" id="{ids()}">\n{ind}\t<DataPath>{path}</DataPath>\n'
    if title:
        s += ml("Title", title, ind + "\t")
    s += extra
    s += f'{ind}\t<ContextMenu name="{name}КонтекстноеМеню" id="{ids()}"/>\n'
    s += f'{ind}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n'
    s += f"{ind}</InputField>\n"
    return s


def popup_menus(ids, ind, list_form=False):
    s = ""
    groups = [("ПодменюСоздатьНаОсновании", "Створити на підставі", True, "CommonPicture.ЗаполнитьПоШаблону", "Picture"),
              ("ПодменюОтчеты", "Звіти", True, "CommonPicture.Отчеты", "PictureAndText")]
    for name, title, sub, pic, rep in groups:
        s += f'{ind}<Popup name="{name}" id="{ids()}">\n' + ml("Title", title, ind + "\t")
        if list_form:
            s += (f"{ind}\t<Picture>\n{ind}\t\t<xr:Ref>{pic}</xr:Ref>\n{ind}\t\t<xr:LoadTransparent>false</xr:LoadTransparent>\n"
                  f"{ind}\t</Picture>\n{ind}\t<Representation>{rep}</Representation>\n")
        s += f'{ind}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n{ind}\t<ChildItems>\n'
        for suf in ("Важное", "Обычное", "СмТакже"):
            s += (f'{ind}\t\t<ButtonGroup name="{name}{suf}" id="{ids()}">\n'
                  f'{ind}\t\t\t<ExtendedTooltip name="{name}{suf}РасширеннаяПодсказка" id="{ids()}"/>\n{ind}\t\t</ButtonGroup>\n')
        s += f"{ind}\t</ChildItems>\n{ind}</Popup>\n"
    return s


def print_popup(ids, ind):
    return (f'{ind}<Popup name="ПодменюПечать" id="{ids()}">\n' + ml("Title", "Друк", ind + "\t") +
            f"{ind}\t<Picture>\n{ind}\t\t<xr:Ref>StdPicture.Print</xr:Ref>\n{ind}\t\t<xr:LoadTransparent>true</xr:LoadTransparent>\n{ind}\t</Picture>\n"
            f'{ind}\t<ExtendedTooltip name="ПодменюПечатьРасширеннаяПодсказка" id="{ids()}"/>\n{ind}</Popup>\n')


def table_additions(ids, tname, ind):
    s = ""
    for suf, typ in (("СтрокаПоиска", "SearchStringRepresentation"), ("СостояниеПросмотра", "ViewStatusRepresentation"),
                     ("УправлениеПоиском", "SearchControl")):
        tag = {"СтрокаПоиска": "SearchStringAddition", "СостояниеПросмотра": "ViewStatusAddition",
               "УправлениеПоиском": "SearchControlAddition"}[suf]
        s += (f'{ind}<{tag} name="{tname}{suf}" id="{ids()}">\n{ind}\t<AdditionSource>\n{ind}\t\t<Item>{tname}</Item>\n'
              f'{ind}\t\t<Type>{typ}</Type>\n{ind}\t</AdditionSource>\n'
              f'{ind}\t<ContextMenu name="{tname}{suf}КонтекстноеМеню" id="{ids()}"/>\n'
              f'{ind}\t<ExtendedTooltip name="{tname}{suf}РасширеннаяПодсказка" id="{ids()}"/>\n{ind}</{tag}>\n')
    return s


def command(cid, name, title, tip, action, pic=None, shortcut=None, modifies=True):
    ind = "\t\t"
    s = f'{ind}<Command name="{name}" id="{cid}">\n' + ml("Title", title, ind + "\t") + ml("ToolTip", tip, ind + "\t")
    if shortcut:
        s += f"{ind}\t<Shortcut>{shortcut}</Shortcut>\n"
    if pic:
        s += f"{ind}\t<Picture>\n{ind}\t\t<xr:Ref>{pic}</xr:Ref>\n{ind}\t\t<xr:LoadTransparent>true</xr:LoadTransparent>\n{ind}\t</Picture>\n"
    s += f"{ind}\t<Action>{action}</Action>\n"
    if modifies:
        s += f"{ind}\t<ModifiesSavedData>true</ModifiesSavedData>\n"
    s += f"{ind}\t<CurrentRowUse>DontUse</CurrentRowUse>\n{ind}</Command>\n"
    return s


def write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"\xef\xbb\xbf" + text.encode("utf-8"))


# ---------------- ФормаДокумента ----------------
def form_document():
    ids = Ids(1)
    s = HEAD + ml("Title", "Потреба на завершення проєкту", "\t")
    s += "\t<CommandSet>\n\t\t<ExcludedCommand>Post</ExcludedCommand>\n\t\t<ExcludedCommand>PostAndClose</ExcludedCommand>\n\t\t<ExcludedCommand>Write</ExcludedCommand>\n\t</CommandSet>\n"
    s += "\t<AutoTime>CurrentOrLast</AutoTime>\n\t<UsePostingMode>Auto</UsePostingMode>\n\t<RepostOnWrite>true</RepostOnWrite>\n"
    s += '\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1">\n\t\t<ChildItems>\n'
    for bname, cmd, rep, default in (("ФормаПровестиИЗакрыть", "ПровестиИЗакрыть", "Text", True),
                                     ("ФормаЗаписать", "ЗаписатьДокумент", "Picture", False),
                                     ("ФормаПровести", "ПровестиДокумент", "Picture", False)):
        s += (f'\t\t\t<Button name="{bname}" id="{ids()}">\n\t\t\t\t<Type>CommandBarButton</Type>\n'
              f"\t\t\t\t<Representation>{rep}</Representation>\n")
        if default:
            s += "\t\t\t\t<DefaultButton>true</DefaultButton>\n"
        s += (f"\t\t\t\t<CommandName>Form.Command.{cmd}</CommandName>\n"
              f'\t\t\t\t<ExtendedTooltip name="{bname}РасширеннаяПодсказка" id="{ids()}"/>\n\t\t\t</Button>\n')
    s += popup_menus(ids, "\t\t\t")
    s += print_popup(ids, "\t\t\t")
    s += "\t\t</ChildItems>\n\t</AutoCommandBar>\n"
    s += "\t<Events>\n\t\t<Event name=\"OnCreateAtServer\">ПриСозданииНаСервере</Event>\n\t</Events>\n"
    s += "\t<ChildItems>\n"
    # шапка
    s += (f'\t\t<UsualGroup name="ГруппаШапка" id="{ids()}">\n\t\t\t<Group>Vertical</Group>\n\t\t\t<Behavior>Usual</Behavior>\n'
          f"\t\t\t<Representation>None</Representation>\n\t\t\t<ShowTitle>false</ShowTitle>\n"
          f'\t\t\t<ExtendedTooltip name="ГруппаШапкаРасширеннаяПодсказка" id="{ids()}"/>\n\t\t\t<ChildItems>\n')
    s += (f'\t\t\t\t<UsualGroup name="ГруппаНомерДата" id="{ids()}">\n\t\t\t\t\t<Group>Horizontal</Group>\n\t\t\t\t\t<Behavior>Usual</Behavior>\n'
          f"\t\t\t\t\t<Representation>None</Representation>\n\t\t\t\t\t<ShowTitle>false</ShowTitle>\n"
          f'\t\t\t\t\t<ExtendedTooltip name="ГруппаНомерДатаРасширеннаяПодсказка" id="{ids()}"/>\n\t\t\t\t\t<ChildItems>\n')
    s += input_field(ids, "Номер", "Объект.Number", "\t\t\t\t\t\t", "Номер")
    s += input_field(ids, "Дата", "Объект.Date", "\t\t\t\t\t\t", "від")
    s += "\t\t\t\t\t</ChildItems>\n\t\t\t\t</UsualGroup>\n"
    s += input_field(ids, "Подразделение", "Объект.Подразделение", "\t\t\t\t", None,
                     "\t\t\t\t\t<DropListButton>true</DropListButton>\n")
    s += (f'\t\t\t\t<UsualGroup name="ГруппаФайл" id="{ids()}">\n\t\t\t\t\t<Group>Horizontal</Group>\n\t\t\t\t\t<Behavior>Usual</Behavior>\n'
          f"\t\t\t\t\t<Representation>None</Representation>\n\t\t\t\t\t<ShowTitle>false</ShowTitle>\n"
          f'\t\t\t\t\t<ExtendedTooltip name="ГруппаФайлРасширеннаяПодсказка" id="{ids()}"/>\n\t\t\t\t\t<ChildItems>\n')
    s += input_field(ids, "ИмяФайла", "Объект.ИмяФайла", "\t\t\t\t\t\t", None, "\t\t\t\t\t\t\t<ReadOnly>true</ReadOnly>\n")
    s += input_field(ids, "РазделФайла", "Объект.РазделФайла", "\t\t\t\t\t\t", None, "\t\t\t\t\t\t\t<ReadOnly>true</ReadOnly>\n")
    s += "\t\t\t\t\t</ChildItems>\n\t\t\t\t</UsualGroup>\n"
    s += (f'\t\t\t\t<CheckBoxField name="ЦеныБезНДС" id="{ids()}">\n\t\t\t\t\t<DataPath>ЦеныБезНДС</DataPath>\n'
          + ml("Title", "Ціни у файлі без ПДВ (при завантаженні +20 %)", "\t\t\t\t\t") +
          "\t\t\t\t\t<TitleLocation>Right</TitleLocation>\n\t\t\t\t\t<CheckBoxType>Auto</CheckBoxType>\n"
          f'\t\t\t\t\t<ContextMenu name="ЦеныБезНДСКонтекстноеМеню" id="{ids()}"/>\n'
          f'\t\t\t\t\t<ExtendedTooltip name="ЦеныБезНДСРасширеннаяПодсказка" id="{ids()}"/>\n\t\t\t\t</CheckBoxField>\n')
    s += "\t\t\t</ChildItems>\n\t\t</UsualGroup>\n"
    # таблица
    t = "Потреба"
    s += (f'\t\t<Table name="{t}" id="{ids()}">\n\t\t\t<Footer>true</Footer>\n\t\t\t<DataPath>Объект.{t}</DataPath>\n'
          "\t\t\t<RowFilter xsi:nil=\"true\"/>\n"
          f'\t\t\t<ContextMenu name="{t}КонтекстноеМеню" id="{ids()}"/>\n'
          f'\t\t\t<AutoCommandBar name="{t}КоманднаяПанель" id="{ids()}">\n\t\t\t\t<ChildItems>\n'
          f'\t\t\t\t\t<Button name="{t}ЗаполнитьИзExcel" id="{ids()}">\n\t\t\t\t\t\t<Type>CommandBarButton</Type>\n'
          "\t\t\t\t\t\t<Representation>PictureAndText</Representation>\n"
          "\t\t\t\t\t\t<CommandName>Form.Command.ЗаполнитьИзExcel</CommandName>\n"
          f'\t\t\t\t\t\t<ExtendedTooltip name="{t}ЗаполнитьИзExcelРасширеннаяПодсказка" id="{ids()}"/>\n\t\t\t\t\t</Button>\n'
          "\t\t\t\t</ChildItems>\n\t\t\t</AutoCommandBar>\n"
          f'\t\t\t<ExtendedTooltip name="{t}РасширеннаяПодсказка" id="{ids()}"/>\n')
    s += table_additions(ids, t, "\t\t\t")
    s += "\t\t\t<ChildItems>\n"
    cols = [("НомерСтроки", "LineNumber", None, ""),
            ("НомерСтрокиФайла", "НомерСтрокиФайла", None, ""),
            ("НазваниеТовара", "НазваниеТовара", None, ""),
            ("НазваниеСС", "НазваниеСС", None, ""),
            ("ОбщееНазвание", "ОбщееНазвание", None, ""),
            ("Количество", "Количество", None, ""),
            ("Единица", "Единица", None, ""),
            ("КоличествоСС", "КоличествоСС", None, ""),
            ("ЕдиницаСС", "ЕдиницаСС", None, ""),
            ("Цена", "Цена", None, ""),
            ("Сумма", "Сумма", None, f"\t\t\t\t\t<FooterDataPath>Объект.{t}.TotalСумма</FooterDataPath>\n"),
            ("Дома", "Дома", None, ""),
            ("Предупреждение", "Предупреждение", None, "")]
    for name, field, title, extra in cols:
        s += input_field(ids, f"{t}{name}", f"Объект.{t}.{field}", "\t\t\t\t", title,
                         "\t\t\t\t\t<EditMode>EnterOnInput</EditMode>\n" + extra)
    s += "\t\t\t</ChildItems>\n\t\t</Table>\n"
    s += input_field(ids, "Ответственный", "Объект.Ответственный", "\t\t")
    s += input_field(ids, "Комментарий", "Объект.Комментарий", "\t\t")
    s += "\t</ChildItems>\n\t<Attributes>\n"
    s += (f'\t\t<Attribute name="Объект" id="1">\n\t\t\t<Type>\n\t\t\t\t<v8:Type>cfg:DocumentObject.{DOC}</v8:Type>\n\t\t\t</Type>\n'
          "\t\t\t<MainAttribute>true</MainAttribute>\n\t\t\t<SavedData>true</SavedData>\n\t\t</Attribute>\n")
    s += ('\t\t<Attribute name="ЦеныБезНДС" id="2">\n' + ml("Title", "Ціни у файлі без ПДВ", "\t\t\t") +
          "\t\t\t<Type>\n\t\t\t\t<v8:Type>xs:boolean</v8:Type>\n\t\t\t</Type>\n\t\t</Attribute>\n")
    s += "\t</Attributes>\n\t<Commands>\n"
    s += command(1, "ПровестиИЗакрыть", "Провести та закрити", "Провести документ і закрити форму", "ПровестиИЗакрыть", None, "Ctrl+Enter")
    s += command(2, "ЗаписатьДокумент", "Записати", "Записати документ", "ЗаписатьДокумент", "StdPicture.Write", "Ctrl+S")
    s += command(3, "ПровестиДокумент", "Провести", "Провести документ", "ПровестиДокумент", "StdPicture.Post")
    s += command(4, "ЗаполнитьИзExcel", "Заповнити з Excel", "Заповнити потребу з файлу директора (.xlsx)", "ЗаполнитьИзExcel",
                 "StdPicture.OpenFile", None, True)
    s += "\t</Commands>\n</Form>"
    write(ROOT / "ФормаДокумента" / "Ext" / "Form.xml", s)


# ---------------- ФормаСписка / ФормаВыбора ----------------
LIST_EXCL = ["Abort", "Cancel", "CancelSearch", "Change", "Copy", "Create", "Delete", "DynamicListStandardSettings", "Find",
             "FindByCurrentValue", "Ignore", "ListSettings", "LoadDynamicListSettings", "No", "OK", "OutputList", "Post",
             "Refresh", "RestoreValues", "Retry", "SaveDynamicListSettings", "SaveValues", "SetDateInterval",
             "SetDeletionMark", "UndoPosting", "Yes"]


def form_list(fname, title, choice=False):
    ids = Ids(1)
    s = HEAD + ml("Title", title, "\t")
    s += "\t<CommandSet>\n" + "".join(f"\t\t<ExcludedCommand>{c}</ExcludedCommand>\n" for c in LIST_EXCL) + "\t</CommandSet>\n"
    s += '\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1">\n\t\t<Autofill>false</Autofill>\n\t</AutoCommandBar>\n'
    s += "\t<Events>\n\t\t<Event name=\"OnCreateAtServer\">ПриСозданииНаСервере</Event>\n\t</Events>\n\t<ChildItems>\n"
    s += (f'\t\t<Table name="Список" id="{ids()}">\n\t\t\t<DataPath>Список</DataPath>\n'
          "\t\t\t<RowPictureDataPath>Список.DefaultPicture</RowPictureDataPath>\n")
    if choice:
        s += "\t\t\t<ChoiceMode>true</ChoiceMode>\n"
    s += ("\t\t\t<AutoRefresh>false</AutoRefresh>\n\t\t\t<AutoRefreshPeriod>60</AutoRefreshPeriod>\n"
          "\t\t\t<Period>\n\t\t\t\t<v8:variant xsi:type=\"v8:StandardPeriodVariant\">Custom</v8:variant>\n"
          "\t\t\t\t<v8:startDate>0001-01-01T00:00:00</v8:startDate>\n\t\t\t\t<v8:endDate>0001-01-01T00:00:00</v8:endDate>\n\t\t\t</Period>\n"
          "\t\t\t<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>\n\t\t\t<RestoreCurrentRow>false</RestoreCurrentRow>\n"
          "\t\t\t<TopLevelParent xsi:nil=\"true\"/>\n\t\t\t<ShowRoot>true</ShowRoot>\n\t\t\t<AllowRootChoice>false</AllowRootChoice>\n"
          "\t\t\t<UpdateOnDataChange>Auto</UpdateOnDataChange>\n"
          f'\t\t\t<ContextMenu name="СписокКонтекстноеМеню" id="{ids()}"/>\n'
          f'\t\t\t<AutoCommandBar name="СписокКоманднаяПанель" id="{ids()}">\n\t\t\t\t<ChildItems>\n')
    s += popup_menus(ids, "\t\t\t\t\t", list_form=True)
    s += (f'\t\t\t\t\t<ButtonGroup name="ГруппаГлобальныеКоманды" id="{ids()}">\n'
          "\t\t\t\t\t\t<CommandSource>FormCommandPanelGlobalCommands</CommandSource>\n"
          f'\t\t\t\t\t\t<ExtendedTooltip name="ГруппаГлобальныеКомандыРасширеннаяПодсказка" id="{ids()}"/>\n\t\t\t\t\t\t<ChildItems>\n')
    s += print_popup(ids, "\t\t\t\t\t\t\t")
    s += ("\t\t\t\t\t\t</ChildItems>\n\t\t\t\t\t</ButtonGroup>\n"
          f'\t\t\t\t\t<ButtonGroup name="ГруппаКомандыФормы" id="{ids()}">\n\t\t\t\t\t\t<CommandSource>Form</CommandSource>\n'
          f'\t\t\t\t\t\t<ExtendedTooltip name="ГруппаКомандыФормыРасширеннаяПодсказка" id="{ids()}"/>\n\t\t\t\t\t</ButtonGroup>\n'
          "\t\t\t\t</ChildItems>\n\t\t\t</AutoCommandBar>\n"
          f'\t\t\t<ExtendedTooltip name="СписокРасширеннаяПодсказка" id="{ids()}"/>\n')
    s += table_additions(ids, "Список", "\t\t\t")
    s += "\t\t\t<Events>\n\t\t\t\t<Event name=\"OnActivateRow\">СписокПриАктивизацииСтроки</Event>\n\t\t\t</Events>\n\t\t\t<ChildItems>\n"
    for name, field, title in (("Номер", "Number", "Номер"), ("Дата", "Date", "Дата"), ("Подразделение", "Подразделение", None),
                               ("СуммаДокумента", "СуммаДокумента", None), ("ИмяФайла", "ИмяФайла", None),
                               ("РазделФайла", "РазделФайла", None), ("Ответственный", "Ответственный", None),
                               ("Комментарий", "Комментарий", None)):
        s += input_field(ids, name, f"Список.{field}", "\t\t\t\t", title)
    s += "\t\t\t</ChildItems>\n\t\t</Table>\n\t</ChildItems>\n\t<Attributes>\n"
    s += ('\t\t<Attribute name="Список" id="1">\n\t\t\t<Type>\n\t\t\t\t<v8:Type>cfg:DynamicList</v8:Type>\n\t\t\t</Type>\n'
          "\t\t\t<MainAttribute>true</MainAttribute>\n\t\t\t<Settings xsi:type=\"DynamicList\">\n"
          "\t\t\t\t<ManualQuery>false</ManualQuery>\n\t\t\t\t<DynamicDataRead>true</DynamicDataRead>\n"
          f"\t\t\t\t<MainTable>Document.{DOC}</MainTable>\n\t\t\t\t<ListSettings/>\n\t\t\t</Settings>\n\t\t</Attribute>\n")
    s += "\t</Attributes>\n</Form>"
    write(ROOT / fname / "Ext" / "Form.xml", s)


form_document()
form_list("ФормаСписка", "Потреби на завершення проєктів")
form_list("ФормаВыбора", "Вибір потреби на завершення проєкту", choice=True)
print("ok")
