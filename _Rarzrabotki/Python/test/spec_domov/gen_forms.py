# -*- coding: utf-8 -*-
"""Формы справочника А_СпецификацииДомов (П1) в дамп WORKTREE.
--etap 1: ФормаЭлемента (страницы «Склад», «Будинки», «Дерево», «Розгорнутий склад», «Порівняння», «Закупівля»,
          «Факт по будинках»; команды этапа 1), ФормаСписка («Еталон для СС», «Перевірити цикли»), ФормаГруппы.
--etap 2: то же с командами этапа 2 + ФормаПодбораКарточки (Р35) + ФормаПерепривязки (§5 шаг 7).
Модули форм — из spec_domov/bsl/forma_*.bsl (накопительно). UUID обёрток форм — uuid5: перегенерация их не меняет
(память raschet_komplektacij_dokument). Порядок свойств полей и таблиц — канон BASEBuh (E1 §3).
Запуск: python gen_forms.py --etap 1|2"""
import argparse
import os
import re
import sys
import uuid

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_bsl import read_src  # noqa: E402
from common import BB_WT, read_text, write_text, zapros  # noqa: E402

CAT = "А_СпецификацииДомов"
NS = uuid.UUID("3f6a1c52-7d0e-4b8a-9c21-5e4d2b7a9f10")
XNS = {"app": "http://v8.1c.ru/8.2/managed-application/core", "cfg": "http://v8.1c.ru/8.1/data/enterprise/current-config",
      "cmi": "http://v8.1c.ru/8.2/managed-application/cmi", "dcscor": "http://v8.1c.ru/8.1/data-composition-system/core",
      "dcsset": "http://v8.1c.ru/8.1/data-composition-system/settings", "ent": "http://v8.1c.ru/8.1/data/enterprise",
      "lf": "http://v8.1c.ru/8.2/managed-application/logform", "style": "http://v8.1c.ru/8.1/data/ui/style",
      "sys": "http://v8.1c.ru/8.1/data/ui/fonts/system", "v8": "http://v8.1c.ru/8.1/data/core", "v8ui": "http://v8.1c.ru/8.1/data/ui",
      "web": "http://v8.1c.ru/8.1/data/ui/colors/web", "win": "http://v8.1c.ru/8.1/data/ui/colors/windows",
      "xen": "http://v8.1c.ru/8.3/xcf/enums", "xpr": "http://v8.1c.ru/8.3/xcf/predef", "xr": "http://v8.1c.ru/8.3/xcf/readable",
      "xs": "http://www.w3.org/2001/XMLSchema", "xsi": "http://www.w3.org/2001/XMLSchema-instance"}


def xmlns(*prefixes):
    """Объявления в каноническом порядке выгрузки платформы — по алфавиту префиксов (образец А_ТехкартаКМД/Forms)."""
    return " ".join(f'xmlns:{p}="{XNS[p]}"' for p in sorted(prefixes))


FORM_HEAD = ('<?xml version="1.0" encoding="UTF-8"?>\n<Form xmlns="http://v8.1c.ru/8.3/xcf/logform" '
             + xmlns("app", "cfg", "dcscor", "dcsset", "ent", "lf", "style", "sys", "v8", "v8ui", "web", "win", "xr", "xs", "xsi")
             + ' version="2.13">\n')
MD_HEAD = ('<?xml version="1.0" encoding="UTF-8"?>\n<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses" '
           + xmlns("app", "cfg", "cmi", "ent", "lf", "style", "sys", "v8", "v8ui", "web", "win", "xen", "xpr", "xr", "xs", "xsi")
           + ' version="2.13">\n')


def uid(key):
    return str(uuid.uuid5(NS, key))


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def ml(tag, ru, uk, t, attrs=""):
    return (f"{t}<{tag}{attrs}>\n{t}\t<v8:item>\n{t}\t\t<v8:lang>ru</v8:lang>\n{t}\t\t<v8:content>{esc(ru)}</v8:content>\n"
            f"{t}\t</v8:item>\n{t}\t<v8:item>\n{t}\t\t<v8:lang>uk</v8:lang>\n{t}\t\t<v8:content>{esc(uk)}</v8:content>\n"
            f"{t}\t</v8:item>\n{t}</{tag}>\n")


class Ids:
    def __init__(self):
        self.n = 1

    def __call__(self):
        self.n += 1
        return self.n - 1


def events_xml(events, t):
    if not events:
        return ""
    return f"{t}<Events>\n" + "".join(f'{t}\t<Event name="{e}">{h}</Event>\n' for e, h in events) + f"{t}</Events>\n"


def field(ids, name, path, t, kind="InputField", title=None, ro=False, visible=True, in_table=False, width=None,
          height=None, stretch=False, multiline=False, footer=None, checkbox_type=False, title_loc=None, events=None):
    """Порядок (по 518 формам BASEBuh): DataPath, Visible, ReadOnly, Title, TitleLocation, EditMode, CheckBoxType,
    FooterDataPath, Width, Height, HorizontalStretch, MultiLine, ContextMenu, ExtendedTooltip, Events."""
    s = f'{t}<{kind} name="{name}" id="{ids()}">\n{t}\t<DataPath>{path}</DataPath>\n'
    if not visible:
        s += f"{t}\t<Visible>false</Visible>\n"
    if ro:
        s += f"{t}\t<ReadOnly>true</ReadOnly>\n"
    if title:
        s += ml("Title", title[0], title[1], t + "\t")
    if title_loc:
        s += f"{t}\t<TitleLocation>{title_loc}</TitleLocation>\n"
    if in_table:
        s += f"{t}\t<EditMode>EnterOnInput</EditMode>\n"
    if checkbox_type:
        s += f"{t}\t<CheckBoxType>Auto</CheckBoxType>\n"
    if footer:
        s += f"{t}\t<FooterDataPath>{footer}</FooterDataPath>\n"
    if width:
        s += f"{t}\t<Width>{width}</Width>\n"
    if height:
        s += f"{t}\t<Height>{height}</Height>\n"
    if stretch:
        s += f"{t}\t<HorizontalStretch>true</HorizontalStretch>\n"
    if multiline:
        s += f"{t}\t<MultiLine>true</MultiLine>\n"
    s += f'{t}\t<ContextMenu name="{name}КонтекстноеМеню" id="{ids()}"/>\n'
    s += f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n'
    return s + events_xml(events, t + "\t") + f"{t}</{kind}>\n"


def group(ids, name, children, t, orient="Horizontal"):
    s = (f'{t}<UsualGroup name="{name}" id="{ids()}">\n{t}\t<Group>{orient}</Group>\n{t}\t<Behavior>Usual</Behavior>\n'
         f"{t}\t<Representation>None</Representation>\n{t}\t<ShowTitle>false</ShowTitle>\n"
         f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n{t}\t<ChildItems>\n')
    s += "".join(c(ids, t + "\t\t") for c in children)
    return s + f"{t}\t</ChildItems>\n{t}</UsualGroup>\n"


def button(ids, name, command, t):
    return (f'{t}<Button name="{name}" id="{ids()}">\n{t}\t<Type>CommandBarButton</Type>\n'
            f"{t}\t<CommandName>Form.Command.{command}</CommandName>\n"
            f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n{t}</Button>\n')


def additions(ids, name, t):
    s = ""
    for tag, suffix, typ in (("SearchStringAddition", "СтрокаПоиска", "SearchStringRepresentation"),
                             ("ViewStatusAddition", "СостояниеПросмотра", "ViewStatusRepresentation"),
                             ("SearchControlAddition", "УправлениеПоиском", "SearchControl")):
        n = name + suffix
        s += (f'{t}<{tag} name="{n}" id="{ids()}">\n{t}\t<AdditionSource>\n{t}\t\t<Item>{name}</Item>\n'
              f"{t}\t\t<Type>{typ}</Type>\n{t}\t</AdditionSource>\n"
              f'{t}\t<ContextMenu name="{n}КонтекстноеМеню" id="{ids()}"/>\n'
              f'{t}\t<ExtendedTooltip name="{n}РасширеннаяПодсказка" id="{ids()}"/>\n{t}</{tag}>\n')
    return s


def table(ids, name, path, columns, t, buttons=(), events=None, ro=False, tree=False, change_rows=True, height=None):
    """Порядок: Representation, ReadOnly, ChangeRowSet, ChangeRowOrder, Height, AutoInsertNewRow, InitialTreeView,
    EnableStartDrag, EnableDrag, DataPath, RowFilter, ContextMenu, AutoCommandBar, ExtendedTooltip, дополнения, Events, ChildItems."""
    s = f'{t}<Table name="{name}" id="{ids()}">\n{t}\t<Representation>{"Tree" if tree else "List"}</Representation>\n'
    if ro:
        s += f"{t}\t<ReadOnly>true</ReadOnly>\n"
    if not change_rows:
        s += f"{t}\t<ChangeRowSet>false</ChangeRowSet>\n{t}\t<ChangeRowOrder>false</ChangeRowOrder>\n"
    if height:
        s += f"{t}\t<Height>{height}</Height>\n"
    if change_rows:
        s += f"{t}\t<AutoInsertNewRow>true</AutoInsertNewRow>\n"
    if tree:
        s += f"{t}\t<InitialTreeView>ExpandAllLevels</InitialTreeView>\n"
    s += f"{t}\t<EnableStartDrag>true</EnableStartDrag>\n{t}\t<EnableDrag>true</EnableDrag>\n{t}\t<DataPath>{path}</DataPath>\n"
    if not tree:
        s += f'{t}\t<RowFilter xsi:nil="true"/>\n'
    s += f'{t}\t<ContextMenu name="{name}КонтекстноеМеню" id="{ids()}"/>\n'
    if buttons:
        s += f'{t}\t<AutoCommandBar name="{name}КоманднаяПанель" id="{ids()}">\n{t}\t\t<ChildItems>\n'
        s += "".join(button(ids, f"{name}{cmd}", cmd, t + "\t\t\t") for cmd in buttons)
        s += f"{t}\t\t</ChildItems>\n{t}\t</AutoCommandBar>\n"
    else:
        s += f'{t}\t<AutoCommandBar name="{name}КоманднаяПанель" id="{ids()}"/>\n'
    s += f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n'
    s += additions(ids, name, t + "\t") + events_xml(events, t + "\t")
    s += f"{t}\t<ChildItems>\n" + "".join(c(ids, t + "\t\t") for c in columns) + f"{t}\t</ChildItems>\n"
    return s + f"{t}</Table>\n"


def dyn_table(ids, name, columns, t, user_group):
    """Канон ОбщиеНазванияНоменклатуры.ФормаСписка (E1 §3.6): панель таблицы скрыта, команды — в панели формы."""
    s = (f'{t}<Table name="{name}" id="{ids()}">\n{t}\t<CommandBarLocation>None</CommandBarLocation>\n'
         f"{t}\t<DefaultItem>true</DefaultItem>\n{t}\t<UseAlternationRowColor>true</UseAlternationRowColor>\n"
         f"{t}\t<InitialTreeView>ExpandTopLevel</InitialTreeView>\n{t}\t<EnableStartDrag>true</EnableStartDrag>\n"
         f"{t}\t<EnableDrag>true</EnableDrag>\n{t}\t<DataPath>{name}</DataPath>\n"
         f"{t}\t<RowPictureDataPath>{name}.DefaultPicture</RowPictureDataPath>\n"
         f"{t}\t<AutoRefresh>false</AutoRefresh>\n{t}\t<AutoRefreshPeriod>60</AutoRefreshPeriod>\n{t}\t<Period>\n"
         f'{t}\t\t<v8:variant xsi:type="v8:StandardPeriodVariant">Custom</v8:variant>\n'
         f"{t}\t\t<v8:startDate>0001-01-01T00:00:00</v8:startDate>\n{t}\t\t<v8:endDate>0001-01-01T00:00:00</v8:endDate>\n"
         f"{t}\t</Period>\n{t}\t<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>\n{t}\t<RestoreCurrentRow>false</RestoreCurrentRow>\n"
         f'{t}\t<TopLevelParent xsi:nil="true"/>\n{t}\t<ShowRoot>true</ShowRoot>\n{t}\t<AllowRootChoice>false</AllowRootChoice>\n'
         f"{t}\t<UpdateOnDataChange>Auto</UpdateOnDataChange>\n{t}\t<UserSettingsGroup>{user_group}</UserSettingsGroup>\n"
         f'{t}\t<ContextMenu name="{name}КонтекстноеМеню" id="{ids()}"/>\n'
         f'{t}\t<AutoCommandBar name="{name}КоманднаяПанель" id="{ids()}">\n{t}\t\t<Autofill>false</Autofill>\n{t}\t</AutoCommandBar>\n'
         f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n')
    s += additions(ids, name, t + "\t")
    s += f"{t}\t<ChildItems>\n" + "".join(c(ids, t + "\t\t") for c in columns) + f"{t}\t</ChildItems>\n"
    return s + f"{t}</Table>\n"


def user_settings_group(ids, name, t):
    return (f'{t}<UsualGroup name="{name}" id="{ids()}">\n' + ml("Title", "Группа пользовательских настроек",
            "Група користувацьких налаштувань", t + "\t") +
            f"{t}\t<VerticalStretch>false</VerticalStretch>\n{t}\t<Group>Vertical</Group>\n{t}\t<Behavior>Usual</Behavior>\n"
            f'{t}\t<ShowTitle>false</ShowTitle>\n{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n'
            f"{t}</UsualGroup>\n")


def page(ids, name, ru, uk, children, t):
    s = f'{t}<Page name="{name}" id="{ids()}">\n' + ml("Title", ru, uk, t + "\t")
    s += f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n{t}\t<ChildItems>\n'
    s += "".join(c(ids, t + "\t\t") for c in children)
    return s + f"{t}\t</ChildItems>\n{t}</Page>\n"


def pages(ids, name, children, t):
    s = (f'{t}<Pages name="{name}" id="{ids()}">\n{t}\t<PagesRepresentation>TabsOnTop</PagesRepresentation>\n'
         f'{t}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{ids()}"/>\n{t}\t<ChildItems>\n')
    s += "".join(c(ids, t + "\t\t") for c in children)
    return s + f"{t}\t</ChildItems>\n{t}</Pages>\n"


def type_xml(kind, t):
    if kind.startswith("cfg:"):
        return f"{t}<v8:Type>{kind}</v8:Type>\n"
    if kind == "bool":
        return f"{t}<v8:Type>xs:boolean</v8:Type>\n"
    if kind == "uuid":
        return f"{t}<v8:Type>v8:UUID</v8:Type>\n"
    if kind == "vt":
        return f"{t}<v8:Type>v8:ValueTable</v8:Type>\n"
    if kind == "vtree":
        return f"{t}<v8:Type>v8:ValueTree</v8:Type>\n"
    if kind.startswith("str:"):
        return (f"{t}<v8:Type>xs:string</v8:Type>\n{t}<v8:StringQualifiers>\n{t}\t<v8:Length>{kind[4:]}</v8:Length>\n"
                f"{t}\t<v8:AllowedLength>Variable</v8:AllowedLength>\n{t}</v8:StringQualifiers>\n")
    if kind.startswith("num:"):
        d, f = kind[4:].split(",")
        return (f"{t}<v8:Type>xs:decimal</v8:Type>\n{t}<v8:NumberQualifiers>\n{t}\t<v8:Digits>{d}</v8:Digits>\n"
                f"{t}\t<v8:FractionDigits>{f}</v8:FractionDigits>\n{t}\t<v8:AllowedSign>Any</v8:AllowedSign>\n{t}</v8:NumberQualifiers>\n")
    raise ValueError(kind)


def attr(aid, name, kind, t, title=None, columns=None, main=False, saved=False):
    s = f'{t}<Attribute name="{name}" id="{aid}">\n'
    if title:
        s += ml("Title", title[0], title[1], t + "\t")
    s += f"{t}\t<Type>\n" + type_xml(kind, t + "\t\t") + f"{t}\t</Type>\n"
    if main:
        s += f"{t}\t<MainAttribute>true</MainAttribute>\n"
    if saved:
        s += f"{t}\t<SavedData>true</SavedData>\n"
    if columns:
        s += f"{t}\t<Columns>\n"
        for cid, (cname, ckind, cru, cuk) in enumerate(columns, 1):
            s += f'{t}\t\t<Column name="{cname}" id="{cid}">\n' + ml("Title", cru, cuk, t + "\t\t\t")
            s += f"{t}\t\t\t<Type>\n" + type_xml(ckind, t + "\t\t\t\t") + f"{t}\t\t\t</Type>\n{t}\t\t</Column>\n"
        s += f"{t}\t</Columns>\n"
    return s + f"{t}</Attribute>\n"


def dyn_attr(aid, query, order_field, t):
    q_xml = query.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return (f'{t}<Attribute name="Список" id="{aid}">\n{t}\t<Type>\n{t}\t\t<v8:Type>cfg:DynamicList</v8:Type>\n{t}\t</Type>\n'
            f"{t}\t<MainAttribute>true</MainAttribute>\n{t}\t<Settings xsi:type=\"DynamicList\">\n"
            f"{t}\t\t<ManualQuery>true</ManualQuery>\n{t}\t\t<DynamicDataRead>true</DynamicDataRead>\n"
            f"{t}\t\t<QueryText>{q_xml}</QueryText>\n{t}\t\t<MainTable>Catalog.{CAT}</MainTable>\n{t}\t\t<ListSettings>\n"
            f"{t}\t\t\t<dcsset:filter>\n{t}\t\t\t\t<dcsset:viewMode>Normal</dcsset:viewMode>\n"
            f"{t}\t\t\t\t<dcsset:userSettingID>dfcece9d-5077-440b-b6b3-45a5cb4538eb</dcsset:userSettingID>\n{t}\t\t\t</dcsset:filter>\n"
            f"{t}\t\t\t<dcsset:order>\n{t}\t\t\t\t<dcsset:item xsi:type=\"dcsset:OrderItemField\">\n"
            f"{t}\t\t\t\t\t<dcsset:field>{order_field}</dcsset:field>\n{t}\t\t\t\t\t<dcsset:orderType>Asc</dcsset:orderType>\n"
            f"{t}\t\t\t\t</dcsset:item>\n{t}\t\t\t\t<dcsset:viewMode>Normal</dcsset:viewMode>\n"
            f"{t}\t\t\t\t<dcsset:userSettingID>88619765-ccb3-46c6-ac52-38e9c992ebd4</dcsset:userSettingID>\n{t}\t\t\t</dcsset:order>\n"
            f"{t}\t\t\t<dcsset:conditionalAppearance>\n{t}\t\t\t\t<dcsset:viewMode>Normal</dcsset:viewMode>\n"
            f"{t}\t\t\t\t<dcsset:userSettingID>b75fecce-942b-4aed-abc9-e6a02e460fb3</dcsset:userSettingID>\n"
            f"{t}\t\t\t</dcsset:conditionalAppearance>\n{t}\t\t\t<dcsset:itemsViewMode>Normal</dcsset:itemsViewMode>\n"
            f"{t}\t\t\t<dcsset:itemsUserSettingID>911b6018-f537-43e8-a417-da56b22f9aec</dcsset:itemsUserSettingID>\n"
            f"{t}\t\t</ListSettings>\n{t}\t</Settings>\n{t}</Attribute>\n")


def command(cid, name, ru, uk, t, modifies=False):
    s = f'{t}<Command name="{name}" id="{cid}">\n' + ml("Title", ru, uk, t + "\t") + ml("ToolTip", ru, uk, t + "\t")
    s += f"{t}\t<Action>{name}</Action>\n"
    if modifies:
        s += f"{t}\t<ModifiesSavedData>true</ModifiesSavedData>\n"
    return s + f"{t}</Command>\n"


def parameter(name, kind, t):
    return f'{t}<Parameter name="{name}">\n{t}\t<Type>\n' + type_xml(kind, t + "\t\t") + f"{t}\t</Type>\n{t}</Parameter>\n"


def form_xml(body, attrs, cmds="", params="", events=(), use_for=None, bar=""):
    s = FORM_HEAD
    if use_for:
        s += f"\t<UseForFoldersAndItems>{use_for}</UseForFoldersAndItems>\n"
    if bar:
        s += '\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1">\n\t\t<ChildItems>\n' + bar + "\t\t</ChildItems>\n\t</AutoCommandBar>\n"
    else:
        s += '\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1"/>\n'
    if events:
        s += "\t<Events>\n" + "".join(f'\t\t<Event name="{e}">{h}</Event>\n' for e, h in events) + "\t</Events>\n"
    s += "\t<ChildItems>\n" + body + "\t</ChildItems>\n\t<Attributes>\n" + attrs + "\t</Attributes>\n"
    if cmds:
        s += "\t<Commands>\n" + cmds + "\t</Commands>\n"
    if params:
        s += "\t<Parameters>\n" + params + "\t</Parameters>\n"
    return s + "</Form>\n"


def form_md(form, ru, uk):
    return (MD_HEAD + f'\t<Form uuid="{uid(CAT + "." + form)}">\n\t\t<Properties>\n\t\t\t<Name>{form}</Name>\n'
            + ml("Synonym", ru, uk, "\t\t\t") +
            "\t\t\t<Comment/>\n\t\t\t<FormType>Managed</FormType>\n\t\t\t<IncludeHelpInContents>false</IncludeHelpInContents>\n"
            "\t\t\t<UsePurposes>\n\t\t\t\t<v8:Value xsi:type=\"app:ApplicationUsePurpose\">PlatformApplication</v8:Value>\n"
            "\t\t\t\t<v8:Value xsi:type=\"app:ApplicationUsePurpose\">MobilePlatformApplication</v8:Value>\n"
            "\t\t\t</UsePurposes>\n\t\t</Properties>\n\t</Form>\n</MetaDataObject>\n")


def register_form(form, default_tag=None):
    path = os.path.join(BB_WT, f"Catalogs/{CAT}.xml")
    s, _, _ = read_text(path)
    s2 = s
    if default_tag:
        ref = f"Catalog.{CAT}.Form.{form}"
        s2 = re.sub(rf"<{default_tag}/>|<{default_tag}>[^<]*</{default_tag}>", f"<{default_tag}>{ref}</{default_tag}>", s2, count=1)
        assert f"<{default_tag}>{ref}</{default_tag}>" in s2, default_tag
    if f"<Form>{form}</Form>" not in s2:
        head, sep, tail = s2.rpartition("\t\t</ChildObjects>\n\t</Catalog>")
        assert sep, "конец ChildObjects справочника"
        s2 = head + f"\t\t\t<Form>{form}</Form>\n" + sep + tail
    if s2 != s:
        write_text(path, s2)


def write_form(form, ru, uk, xml, module_parts, default_tag=None):
    base = f"Catalogs/{CAT}/Forms/{form}"
    write_text(os.path.join(BB_WT, base + ".xml"), form_md(form, ru, uk))
    write_text(os.path.join(BB_WT, base + "/Ext/Form.xml"), xml)
    if module_parts:
        write_text(os.path.join(BB_WT, base + "/Ext/Form/Module.bsl"), "\n".join(read_src(p).rstrip("\n") + "\n" for p in module_parts))
    register_form(form, default_tag)
    print("  форма", form)


def F(name, path, **kw):
    return lambda i, t: field(i, name, path, t, **kw)


def C(name, path, **kw):
    """Колонка таблицы формы."""
    return lambda i, t: field(i, name, path, t, in_table=True, **kw)


def G(name, children, **kw):
    return lambda i, t: group(i, name, children, t, **kw)


def T(name, path, columns, **kw):
    return lambda i, t: table(i, name, path, columns, t, **kw)


def P(name, ru, uk, children):
    return lambda i, t: page(i, name, ru, uk, children, t)


REF_UNIT = "cfg:CatalogRef.КлассификаторЕдиницИзмерения"
TREE_COLS = [("Представление", "str:250", "Строка", "Рядок"), ("ПозицияСС", "str:200", "Позиция СС", "Позиція СС"),
             ("Количество", "num:15,6", "Количество на дом", "Кількість на будинок"),
             ("ЕдиницаСтроки", REF_UNIT, "Единица строки", "Одиниця рядка"),
             ("КоличествоКарточки", "num:15,6", "В ед. карточки", "В од. картки"),
             ("СуммаНормы", "num:15,2", "Сумма нормы, с НДС", "Сума норми, з ПДВ")]
RAZV_COLS = [("Номенклатура", "cfg:CatalogRef.Номенклатура", "Карточка", "Картка"),
             ("ОбщееНазвание", "cfg:CatalogRef.ОбщиеНазванияНоменклатуры", "Общее название", "Загальна назва"),
             ("ПозицияСС", "str:200", "Позиция СС", "Позиція СС"),
             ("КоличествоЕдСтроки", "num:15,6", "Количество в ед. строки", "Кількість в од. рядка"),
             ("ЕдиницаСтроки", REF_UNIT, "Единица строки", "Одиниця рядка"),
             ("Количество", "num:15,6", "Норма в ед. карточки", "Норма в од. картки"),
             ("ЦенаНормы", "num:16,6", "Цена нормы, с НДС", "Ціна норми, з ПДВ"),
             ("СуммаНормы", "num:15,2", "Сумма нормы, с НДС", "Сума норми, з ПДВ"),
             ("ПутьУзла", "str:200", "Путь узлов", "Шлях вузлів"), ("НеПоставляется", "bool", "Не поставляется", "Не постачається")]
SR_COLS = [("НомерСтроки", "num:5,0", "№ СС", "№ СС"), ("ПозицияСС", "str:200", "Позиция СС", "Позиція СС"),
           ("КоличествоСС", "num:15,6", "Количество СС", "Кількість СС"), ("ЕдиницаСС", "str:20", "Ед. СС", "Од. СС"),
           ("КоличествоСпец", "num:15,6", "Количество спецификации", "Кількість специфікації"),
           ("СуммаСС", "num:15,2", "Сумма СС", "Сума СС"), ("СуммаНормы", "num:15,2", "Сумма нормы", "Сума норми"),
           ("Статус", "str:100", "Статус", "Статус")]
CMD1 = [("СделатьЭталономСС", "Сделать эталоном для СС", "Зробити еталоном для СС", False),
        ("ПоказатьВходимость", "Входимость", "Входимість", False),
        ("ОбновитьДерево", "Обновить дерево", "Оновити дерево", False),
        ("РазвернутьДоДеталей", "Развернуть до деталей", "Розгорнути до деталей", False)]
CMD2_SOSTAV = [("ЗаполнитьИзСС", "Заполнить из СС", "Заповнити з СС", True),
               ("ОтметитьЗеленыеПроверенными", "Отметить зелёные проверенными", "Позначити зелені перевіреними", True),
               ("ПодобратьКарточку", "Подобрать карточку", "Підібрати картку", True),
               ("ВзятьИзОригиналаExcel", "Взять из оригинала Excel", "Взяти з оригіналу Excel", True),
               ("ОбъединитьВУзел", "Объединить в узел", "Обʼєднати у вузол", True),
               ("РасформироватьУзел", "Расформировать узел", "Розформувати вузол", True),
               ("СоставИзДокументаКомплектации", "Состав из документа комплектации", "Склад з документа комплектації", True)]
CMD2_FORM = [("РассчитатьЦеныНормы", "Рассчитать цены нормы", "Розрахувати ціни норми", True),
             ("СравнитьСоСС", "Сравнить с СС", "Порівняти з СС", False),
             ("НулевыеПозицииСФактом", "Нулевые позиции с фактом", "Нульові позиції з фактом", False),
             ("СкопироватьИзДругойСпецификации", "Скопировать карточки и узлы из другой спецификации",
              "Скопіювати картки й дерево вузлів з іншої специфікації", True)]


def forma_elementa(etap):
    ids = Ids()
    sostav_cols = [
        C("СоставНомерСтроки", "Объект.Состав.LineNumber", kind="LabelField", width=4),
        C("СоставПроверено", "Объект.Состав.Проверено", kind="CheckBoxField", checkbox_type=True),
        C("СоставВидСтроки", "Объект.Состав.ВидСтроки"),
        C("СоставЭтапСС", "Объект.Состав.ЭтапСС"),
        C("СоставПозицияСС", "Объект.Состав.ПозицияСС", width=25),
        C("СоставНоменклатура", "Объект.Состав.Номенклатура", width=25, events=[("OnChange", "СоставНоменклатураПриИзменении")]),
        C("СоставИсточникКарточки", "Объект.Состав.ИсточникКарточки"),
        C("СоставУзел", "Объект.Состав.Узел", events=[("OnChange", "СоставУзелПриИзменении")]),
        C("СоставОбщееНазвание", "Объект.Состав.ОбщееНазвание"),
        C("СоставКоличество", "Объект.Состав.Количество", events=[("OnChange", "СоставКоличествоПриИзменении")]),
        C("СоставЕдиницаСтроки", "Объект.Состав.ЕдиницаСтроки", events=[("OnChange", "СоставЕдиницаСтрокиПриИзменении")]),
        C("СоставКоэффициент", "Объект.Состав.Коэффициент", events=[("OnChange", "СоставКоэффициентПриИзменении")]),
        C("СоставИсточникКоэффициента", "Объект.Состав.ИсточникКоэффициента"),
        C("СоставЦенаНормы", "Объект.Состав.ЦенаНормы"),
        C("СоставНеПоставляется", "Объект.Состав.НеПоставляется", kind="CheckBoxField", checkbox_type=True),
        C("СоставПокупаетсяГотовым", "Объект.Состав.ПокупаетсяГотовым", kind="CheckBoxField", checkbox_type=True),
        C("СоставКомментарий", "Объект.Состав.Комментарий", width=20),
        C("СоставКлючСтроки", "Объект.Состав.КлючСтроки", visible=False),
    ]
    zak_cols = [C(f"ЗакупочныйСостав{n}", f"Объект.ЗакупочныйСостав.{n}") for n in (
        "Номенклатура", "ГруппаЗакупки", "ОбщееНазвание", "Узел", "ЭтапСС", "ПозицияСС", "НормаЭталона", "СреднееФакта",
        "ДомовСФактом", "КоэффициентВариации", "Количество", "ЦенаБезНДС", "ЦенаСНДС", "ИсточникЦены", "ДокументЦены",
        "ЦенаНормы", "Комментарий")]
    zak_cols += [C(f"ЗакупочныйСостав{n}", f"Объект.ЗакупочныйСостав.{n}", kind="CheckBoxField", checkbox_type=True)
                 for n in ("ИзмененоВручную", "ПотребуетРешения", "Проверено")]
    fakt_cols = [C(f"ФактПоДомам{n}", f"Объект.ФактПоДомам.{n}") for n in (
        "Подразделение", "Номенклатура", "Списано", "СписаноСумма", "Лежит", "ЛежитСумма")]
    fakt_cols.append(C("ФактПоДомамИзСборки", "Объект.ФактПоДомам.ИзСборки", kind="CheckBoxField", checkbox_type=True))
    head = G("ГруппаШапка", [
        G("ГруппаШапкаЛево", [
            F("Наименование", "Объект.Description", stretch=True),
            G("ГруппаВидСтатус", [
                F("ВидСпецификации", "Объект.ВидСпецификации", events=[("OnChange", "ВидСпецификацииПриИзменении")]),
                F("Статус", "Объект.Статус"),
                F("ДатаНачала", "Объект.ДатаНачала")]),
            F("СтруктураСебестоимости", "Объект.СтруктураСебестоимости", stretch=True),
            F("Основание", "Объект.Основание", ro=True, stretch=True)], orient="Vertical"),
        G("ГруппаШапкаПраво", [
            G("ГруппаКодРодитель", [F("Код", "Объект.Code"), F("Родитель", "Объект.Parent", title=("Папка", "Папка"))]),
            G("ГруппаВыпуск", [F("КоличествоВыпуска", "Объект.КоличествоВыпуска"), F("ЕдиницаВыпуска", "Объект.ЕдиницаВыпуска")]),
            F("НоменклатураВузла", "Объект.НоменклатураВузла", stretch=True),
            G("ГруппаДаты", [F("ДатаИзменения", "Объект.ДатаИзменения", ro=True), F("ДатаЦенНормы", "Объект.ДатаЦенНормы", ro=True),
                             F("ДатаФакта", "Объект.ДатаФакта", ro=True)])], orient="Vertical")])
    pg = [
        P("СтраницаСостав", "Состав", "Склад", [
            G("ГруппаОтбор", [
                F("ТолькоПроблемные", "ТолькоПроблемные", kind="CheckBoxField", title_loc="Right", checkbox_type=True,
                  events=[("OnChange", "ТолькоПроблемныеПриИзменении")]),
                F("ИтогСравнения", "ИтогСравнения", kind="LabelField", stretch=True)]),
            T("Состав", "Объект.Состав", sostav_cols, buttons=[c[0] for c in CMD2_SOSTAV] if etap == 2 else [])]),
        P("СтраницаДома", "Дома", "Будинки", [T("Дома", "Объект.Дома", [
            C("ДомаНомерСтроки", "Объект.Дома.LineNumber", kind="LabelField", width=4),
            C("ДомаПодразделение", "Объект.Дома.Подразделение"), C("ДомаКомментарий", "Объект.Дома.Комментарий")])]),
        P("СтраницаДерево", "Дерево", "Дерево", [T("ДеревоУзлов", "ДеревоУзлов",
            [C(f"ДеревоУзлов{n}", f"ДеревоУзлов.{n}") for n, *_ in TREE_COLS], tree=True, ro=True, change_rows=False,
            buttons=["ОбновитьДерево"])]),
        P("СтраницаРазвернутыйСостав", "Развёрнутый состав", "Розгорнутий склад", [T("РазвернутыйСостав", "РазвернутыйСостав",
            [C(f"РазвернутыйСостав{n}", f"РазвернутыйСостав.{n}") for n, *_ in RAZV_COLS], ro=True, change_rows=False,
            buttons=["РазвернутьДоДеталей"])]),
        P("СтраницаСравнение", "Сравнение с СС", "Порівняння з СС", [T("СравнениеСС", "СравнениеСС",
            [C(f"СравнениеСС{n}", f"СравнениеСС.{n}") for n, *_ in SR_COLS], ro=True, change_rows=False,
            buttons=["СравнитьСоСС"] if etap == 2 else [])]),
        P("СтраницаЗакупка", "Закупка", "Закупівля", [
            T("ДомаФакта", "Объект.ДомаФакта", [C("ДомаФактаПодразделение", "Объект.ДомаФакта.Подразделение"),
                                                C("ДомаФактаФактСумма", "Объект.ДомаФакта.ФактСумма"),
                                                C("ДомаФактаКомментарий", "Объект.ДомаФакта.Комментарий")], height=4),
            T("ЗакупочныйСостав", "Объект.ЗакупочныйСостав", zak_cols)]),
        P("СтраницаФактПоДомам", "Факт по домам", "Факт по будинках", [T("ФактПоДомам", "Объект.ФактПоДомам", fakt_cols)]),
    ]
    body = head(ids, "\t\t") + pages(ids, "Страницы", pg, "\t\t")
    body += field(ids, "Комментарий", "Объект.Комментарий", "\t\t", height=2, stretch=True, multiline=True)
    cmds = CMD1 + (CMD2_SOSTAV + CMD2_FORM if etap == 2 else [])
    bar_cmds = ["СделатьЭталономСС", "ПоказатьВходимость"] + ([c[0] for c in CMD2_FORM if c[0] != "СравнитьСоСС"] if etap == 2 else [])
    bar = "".join(button(ids, f"Форма{c}", c, "\t\t\t") for c in bar_cmds)
    attrs = ('\t\t<Attribute name="Объект" id="1">\n\t\t\t<Type>\n\t\t\t\t<v8:Type>cfg:CatalogObject.А_СпецификацииДомов</v8:Type>\n'
             "\t\t\t</Type>\n\t\t\t<MainAttribute>true</MainAttribute>\n\t\t\t<SavedData>true</SavedData>\n\t\t</Attribute>\n")
    attrs += attr(2, "ДеревоУзлов", "vtree", "\t\t", title=("Дерево узлов", "Дерево вузлів"), columns=TREE_COLS)
    attrs += attr(3, "РазвернутыйСостав", "vt", "\t\t", title=("Развёрнутый состав", "Розгорнутий склад"), columns=RAZV_COLS)
    attrs += attr(4, "СравнениеСС", "vt", "\t\t", title=("Сравнение с СС", "Порівняння з СС"), columns=SR_COLS)
    attrs += attr(5, "ТолькоПроблемные", "bool", "\t\t", title=("Только проблемные (не проверенные)", "Лише проблемні (не перевірені)"))
    attrs += attr(6, "ИтогСравнения", "str:300", "\t\t", title=("Σ нормы / Σ СС", "Σ норми / Σ СС"))
    cmd_xml = "".join(command(i, n, ru, uk, "\t\t", m) for i, (n, ru, uk, m) in enumerate(cmds, 1))
    xml = form_xml(body, attrs, cmd_xml, events=[("OnCreateAtServer", "ПриСозданииНаСервере")], use_for="Items", bar=bar)
    parts = ["forma_elementa_1.bsl"] + (["forma_elementa_2.bsl"] if etap == 2 else [])
    write_form("ФормаЭлемента", "Форма элемента", "Форма елемента", xml, parts, "DefaultObjectForm")


def forma_spiska(etap):
    ids = Ids()
    cols = [F("Наименование", "Список.Наименование", kind="LabelField"), F("Код", "Список.Код", kind="LabelField"),
            F("ВидСпецификации", "Список.ВидСпецификации", kind="LabelField"),
            F("Статус", "Список.Статус", kind="LabelField"),
            F("СтруктураСебестоимости", "Список.СтруктураСебестоимости", kind="LabelField"),
            F("ЭталонДляСС", "Список.ЭталонДляСС", kind="LabelField", title=("Эталон для СС", "Еталон для СС")),
            F("ДатаИзменения", "Список.ДатаИзменения", kind="LabelField")]
    ug = "СписокКомпоновщикНастроекПользовательскиеНастройки"
    body = user_settings_group(ids, ug, "\t\t") + dyn_table(ids, "Список", cols, "\t\t", ug)
    cmds = [("ПроверитьЦиклы", "Проверить циклы", "Перевірити цикли", False)]
    if etap == 2:
        cmds.append(("ПерепривязатьПозицииСС", "Перепривязать позиции СС…", "Перепривʼязати позиції СС…", False))
    bar = "".join(button(ids, f"Форма{c[0]}", c[0], "\t\t\t") for c in cmds)
    xml = form_xml(body, dyn_attr(1, zapros("spisok_formy"), "Наименование", "\t\t"),
                   "".join(command(i, n, ru, uk, "\t\t", m) for i, (n, ru, uk, m) in enumerate(cmds, 1)), bar=bar)
    parts = ["forma_spiska_1.bsl"] + (["forma_spiska_2.bsl"] if etap == 2 else [])
    write_form("ФормаСписка", "Форма списка", "Форма списку", xml, parts, "DefaultListForm")


def forma_gruppy():
    ids = Ids()
    body = group(ids, "ГруппаНаименованиеКод", [F("Наименование", "Объект.Description", stretch=True),
                                                F("Код", "Объект.Code")], "\t\t")
    body += field(ids, "Родитель", "Объект.Parent", "\t\t", title=("Входит в папку", "Входить у папку"))
    attrs = ('\t\t<Attribute name="Объект" id="1">\n\t\t\t<Type>\n\t\t\t\t<v8:Type>cfg:CatalogObject.А_СпецификацииДомов</v8:Type>\n'
             "\t\t\t</Type>\n\t\t\t<MainAttribute>true</MainAttribute>\n\t\t\t<SavedData>true</SavedData>\n\t\t</Attribute>\n")
    write_form("ФормаГруппы", "Форма группы", "Форма групи", form_xml(body, attrs, use_for="Folders"), [], "DefaultFolderForm")


PODBOR_COLS = [("Номенклатура", "cfg:CatalogRef.Номенклатура", "Карточка", "Картка"), ("Код", "str:20", "Код", "Код"),
               ("Единица", REF_UNIT, "Единица", "Одиниця"),
               ("ОбщееНазвание", "cfg:CatalogRef.ОбщиеНазванияНоменклатуры", "Общее название", "Загальна назва"),
               ("СписаноКоличество", "num:15,3", "Списано на дома типа", "Списано на будинки типу"),
               ("СписаноСумма", "num:15,2", "Списано, без НДС", "Списано, без ПДВ"),
               ("ДомовСФактом", "num:3,0", "Домов", "Будинків"), ("ФактПоДомам", "str:200", "По домам", "По будинках"),
               ("СкладПроекта", "num:15,3", "Склад проекта", "Склад проєкту"),
               ("ЦенаПоследнейЗакупки", "num:15,2", "Цена посл. закупки, с НДС", "Ціна ост. закупівлі, з ПДВ"),
               ("РазмерСовпал", "bool", "Размер из названия совпал", "Розмір з назви збігся")]
PEREPRIV_COLS = [("Применить", "bool", "Применить", "Застосувати"),
                 ("Спецификация", "cfg:CatalogRef.А_СпецификацииДомов", "Спецификация", "Специфікація"),
                 ("КлючСтроки", "uuid", "Ключ строки", "Ключ рядка"),
                 ("СтарыйЭтап", "cfg:CatalogRef.ЭтапыРабот", "Этап (был)", "Етап (був)"),
                 ("СтараяПозиция", "str:200", "Позиция (была)", "Позиція (була)"),
                 ("НовыйЭтап", "cfg:CatalogRef.ЭтапыРабот", "Этап (новый)", "Етап (новий)"),
                 ("НоваяПозиция", "str:200", "Позиция (новая)", "Позиція (нова)"),
                 ("НомерСтрокиСС", "num:5,0", "№ СС", "№ СС"), ("Цвет", "str:10", "Предложение", "Пропозиція"),
                 ("Причина", "str:250", "Причина", "Причина"), ("ДругаяСС", "bool", "Узел другой СС", "Вузол іншої СС")]


def forma_podbora():
    ids = Ids()
    body = group(ids, "ГруппаКонтекст", [F("ОбщееНазвание", "ОбщееНазвание", ro=True), F("ПозицияСС", "ПозицияСС", ro=True, stretch=True)], "\t\t")
    body += group(ids, "ГруппаОтбор", [
        F("ТолькоОН", "ТолькоОН", kind="CheckBoxField", title_loc="Right", checkbox_type=True, events=[("OnChange", "ТолькоОНПриИзменении")]),
        F("СтрокаПоиска", "СтрокаПоиска", stretch=True, events=[("OnChange", "СтрокаПоискаПриИзменении")])], "\t\t")
    body += table(ids, "Кандидаты", "Кандидаты", [C(f"Кандидаты{n}", f"Кандидаты.{n}", **({"kind": "CheckBoxField", "checkbox_type": True} if k == "bool" else {}))
                                                  for n, k, *_ in PODBOR_COLS], "\t\t", ro=True, change_rows=False,
                  events=[("Selection", "КандидатыВыбор")])
    attrs = (attr(1, "Кандидаты", "vt", "\t\t", title=("Кандидаты", "Кандидати"), columns=PODBOR_COLS)
             + attr(2, "ТолькоОН", "bool", "\t\t", title=("Только общее название позиции", "Лише загальна назва позиції"))
             + attr(3, "СтрокаПоиска", "str:100", "\t\t", title=("Поиск в названии", "Пошук у назві"))
             + attr(4, "ТипБудинкуСС", "cfg:CatalogRef.СтруктураСебестоимости", "\t\t", title=("Тип дома (СС)", "Тип будинку (СС)"))
             + attr(5, "ОбщееНазвание", "cfg:CatalogRef.ОбщиеНазванияНоменклатуры", "\t\t", title=("Общее название", "Загальна назва"))
             + attr(6, "ПозицияСС", "str:200", "\t\t", title=("Позиция СС", "Позиція СС")))
    cmds = [("ВыбратьКарточку", "Выбрать", "Вибрати", False), ("ОбновитьКандидатов", "Обновить", "Оновити", False)]
    bar = "".join(button(ids, f"Форма{c[0]}", c[0], "\t\t\t") for c in cmds)
    params = (parameter("ТипБудинкуСС", "cfg:CatalogRef.СтруктураСебестоимости", "\t\t")
              + parameter("ОбщееНазвание", "cfg:CatalogRef.ОбщиеНазванияНоменклатуры", "\t\t") + parameter("ПозицияСС", "str:200", "\t\t"))
    xml = form_xml(body, attrs, "".join(command(i, n, ru, uk, "\t\t", m) for i, (n, ru, uk, m) in enumerate(cmds, 1)),
                   params, events=[("OnCreateAtServer", "ПриСозданииНаСервере")], bar=bar)
    write_form("ФормаПодбораКарточки", "Подбор карточки", "Підбір картки", xml, ["forma_podbora.bsl"])


def forma_perepriv():
    ids = Ids()
    body = field(ids, "ТипБудинкуСС", "ТипБудинкуСС", "\t\t", events=[("OnChange", "ТипБудинкуССПриИзменении")])
    cols = []
    for n, k, *_ in PEREPRIV_COLS:
        kw = {"kind": "CheckBoxField", "checkbox_type": True} if k == "bool" else {}
        if n != "Применить":
            kw["ro"] = True
        if n == "КлючСтроки":
            kw["visible"] = False
        cols.append(C(f"Предложения{n}", f"Предложения.{n}", **kw))
    body += table(ids, "Предложения", "Предложения", cols, "\t\t", change_rows=False)
    attrs = (attr(1, "ТипБудинкуСС", "cfg:CatalogRef.СтруктураСебестоимости", "\t\t", title=("Тип дома (СС)", "Тип будинку (СС)"))
             + attr(2, "Предложения", "vt", "\t\t", title=("Предложения", "Пропозиції"), columns=PEREPRIV_COLS))
    cmds = [("НайтиПредложения", "Найти сироты и предложения", "Знайти сироти й пропозиції", False),
            ("ПрименитьОтмеченные", "Применить отмеченные", "Застосувати позначені", False)]
    bar = "".join(button(ids, f"Форма{c[0]}", c[0], "\t\t\t") for c in cmds)
    xml = form_xml(body, attrs, "".join(command(i, n, ru, uk, "\t\t", m) for i, (n, ru, uk, m) in enumerate(cmds, 1)),
                   parameter("ТипБудинкуСС", "cfg:CatalogRef.СтруктураСебестоимости", "\t\t"),
                   events=[("OnCreateAtServer", "ПриСозданииНаСервере")], bar=bar)
    write_form("ФормаПерепривязки", "Перепривязка позиций СС", "Перепривʼязка позицій СС", xml, ["forma_perepriv.bsl"])


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--etap", type=int, choices=[1, 2], required=True)
    e = ap.parse_args().etap
    forma_elementa(e)
    forma_spiska(e)
    if e == 1:
        forma_gruppy()             # на этапе 2 не меняется: перегенерация затёрла бы канон платформы (Task 28) в worktree
    else:
        forma_podbora()
        forma_perepriv()
    print("OK")
