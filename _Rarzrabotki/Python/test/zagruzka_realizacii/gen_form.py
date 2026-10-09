# -*- coding: utf-8 -*-
"""Form.xml обробки ЗагрузкаТЧРеализацияТоваровУслуг — дизайн 09.10 у стилі робочих місць (А_РозрахункиПДВ...):
командна панель з картинками, шапка, рядок-резюме, закладки «Зіставлення» / «Налаштування» / «Інструкція»,
перемикач-фільтр і легенда кольорів, головна таблиця (єдиний розтягуваний елемент), підсумки (файл / готово / не
підтверджено + контроль з рядком «Разом»), варіанти для рядка (5 рядків). Усі заголовки — українською (ru і uk).
Запуск: python gen_form.py <каталог Загрузка ТЧРеализацияТоваровУслуг>"""
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
КАТАЛОГ = sys.argv[1]
ИМЯ = "ЗагрузкаТЧРеализацияТоваровУслуг"
NS_FORM = ('xmlns="http://v8.1c.ru/8.3/xcf/logform" xmlns:app="http://v8.1c.ru/8.2/managed-application/core" '
           'xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" xmlns:dcscor="http://v8.1c.ru/8.1/data-composition-system/core" '
           'xmlns:dcssch="http://v8.1c.ru/8.1/data-composition-system/schema" xmlns:dcsset="http://v8.1c.ru/8.1/data-composition-system/settings" '
           'xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" '
           'xmlns:style="http://v8.1c.ru/8.1/data/ui/style" xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" '
           'xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" '
           'xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" '
           'xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" xmlns:xs="http://www.w3.org/2001/XMLSchema" '
           'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13"')


class Ід:
    n = 0

    def __call__(self):
        self.n += 1
        return self.n


ід = Ід()


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def текст(тег, т, t, attrs=""):
    """Багатомовний текст (ru і uk — однаково українською: інтерфейс бази український)."""
    т = esc(т)
    return (f"{t}<{тег}{attrs}>\n" + "".join(
        f"{t}\t<v8:item>\n{t}\t\t<v8:lang>{л}</v8:lang>\n{t}\t\t<v8:content>{т}</v8:content>\n{t}\t</v8:item>\n"
        for л in ("ru", "uk")) + f"{t}</{тег}>\n")


def хвіст(імя, t):
    return (f"{t}<ContextMenu name=\"{імя}КонтекстноеМеню\" id=\"{ід()}\"/>\n"
            f"{t}<ExtendedTooltip name=\"{імя}РасширеннаяПодсказка\" id=\"{ід()}\"/>\n")


def події(пари, t):
    if not пари:
        return ""
    return f"{t}<Events>\n" + "".join(f"{t}\t<Event name=\"{п}\">{о}</Event>\n" for п, о in пари) + f"{t}</Events>\n"


def картинка(імя, t):
    return (f"{t}<Picture>\n{t}\t<xr:Ref>StdPicture.{імя}</xr:Ref>\n{t}\t<xr:LoadTransparent>true</xr:LoadTransparent>\n"
            f"{t}</Picture>\n")


def b(v):
    return "true" if v else "false"


def поле(імя, шлях, t, title=None, ro=False, edit_mode=None, title_loc=None, tooltip=None, width=None, amw=None,
         hstretch=None, choice=False, textedit=None, hint=None, події_=None):
    """InputField: DataPath → Title → ReadOnly → EditMode → TitleLocation → ToolTip → Width → AutoMaxWidth →
    HorizontalStretch → ChoiceButton → TextEdit → InputHint → ContextMenu → ExtendedTooltip → Events."""
    s = f"{t}<InputField name=\"{імя}\" id=\"{ід()}\">\n{t}\t<DataPath>{шлях}</DataPath>\n"
    if title:
        s += текст("Title", title, t + "\t")
    if ro:
        s += f"{t}\t<ReadOnly>true</ReadOnly>\n"
    if edit_mode:
        s += f"{t}\t<EditMode>{edit_mode}</EditMode>\n"
    if title_loc:
        s += f"{t}\t<TitleLocation>{title_loc}</TitleLocation>\n"
    if tooltip:
        s += текст("ToolTip", tooltip, t + "\t")
    if width:
        s += f"{t}\t<Width>{width}</Width>\n"
    if amw is not None:
        s += f"{t}\t<AutoMaxWidth>{b(amw)}</AutoMaxWidth>\n"
    if hstretch is not None:
        s += f"{t}\t<HorizontalStretch>{b(hstretch)}</HorizontalStretch>\n"
    if choice:
        s += f"{t}\t<ChoiceButton>true</ChoiceButton>\n"
    if textedit is not None:
        s += f"{t}\t<TextEdit>{b(textedit)}</TextEdit>\n"
    if hint:
        s += текст("InputHint", hint, t + "\t")
    return s + хвіст(імя, t + "\t") + події(події_, t + "\t") + f"{t}</InputField>\n"


def прапорець(імя, шлях, t, title=None, title_loc=None, edit_mode=None, tooltip=None, події_=None):
    """CheckBoxField: DataPath → Title → TitleLocation → ToolTip → EditMode → CheckBoxType."""
    s = f"{t}<CheckBoxField name=\"{імя}\" id=\"{ід()}\">\n{t}\t<DataPath>{шлях}</DataPath>\n"
    if title:
        s += текст("Title", title, t + "\t")
    if title_loc:
        s += f"{t}\t<TitleLocation>{title_loc}</TitleLocation>\n"
    if tooltip:
        s += текст("ToolTip", tooltip, t + "\t")
    if edit_mode:
        s += f"{t}\t<EditMode>{edit_mode}</EditMode>\n"
    s += f"{t}\t<CheckBoxType>Auto</CheckBoxType>\n" + хвіст(імя, t + "\t") + події(події_, t + "\t")
    return s + f"{t}</CheckBoxField>\n"


def напис_поле(імя, шлях, t, hstretch=True):
    """LabelField: DataPath → TitleLocation → AutoMaxWidth → HorizontalStretch."""
    return (f"{t}<LabelField name=\"{імя}\" id=\"{ід()}\">\n{t}\t<DataPath>{шлях}</DataPath>\n"
            f"{t}\t<TitleLocation>None</TitleLocation>\n{t}\t<AutoMaxWidth>false</AutoMaxWidth>\n"
            f"{t}\t<HorizontalStretch>{b(hstretch)}</HorizontalStretch>\n" + хвіст(імя, t + "\t") + f"{t}</LabelField>\n")


def напис(імя, т, t):
    """LabelDecoration: Title (formatted=false)."""
    return (f"{t}<LabelDecoration name=\"{імя}\" id=\"{ід()}\">\n" + текст("Title", т, t + "\t", ' formatted="false"')
            + хвіст(імя, t + "\t") + f"{t}</LabelDecoration>\n")


def група(імя, t, вміст, гориз=True, title=None, repr_="None", ширина_дітей=None):
    """UsualGroup: Title → Group → Behavior → Representation → ChildItemsWidth → ShowTitle."""
    s = f"{t}<UsualGroup name=\"{імя}\" id=\"{ід()}\">\n"
    if title:
        s += текст("Title", title, t + "\t")
    s += (f"{t}\t<Group>{'Horizontal' if гориз else 'Vertical'}</Group>\n{t}\t<Behavior>Usual</Behavior>\n"
          f"{t}\t<Representation>{repr_}</Representation>\n")
    if ширина_дітей:
        s += f"{t}\t<ChildItemsWidth>{ширина_дітей}</ChildItemsWidth>\n"
    s += (f"{t}\t<ShowTitle>{b(title)}</ShowTitle>\n"
          f"{t}\t<ExtendedTooltip name=\"{імя}РасширеннаяПодсказка\" id=\"{ід()}\"/>\n{t}\t<ChildItems>\n")
    return s + вміст(t + "\t\t") + f"{t}\t</ChildItems>\n{t}</UsualGroup>\n"


def сторінка(імя, title, pic, t, вміст):
    """Page: Title → Picture → ExtendedTooltip → ChildItems."""
    return (f"{t}<Page name=\"{імя}\" id=\"{ід()}\">\n" + текст("Title", title, t + "\t") + картинка(pic, t + "\t")
            + f"{t}\t<ExtendedTooltip name=\"{імя}РасширеннаяПодсказка\" id=\"{ід()}\"/>\n{t}\t<ChildItems>\n"
            + вміст(t + "\t\t") + f"{t}\t</ChildItems>\n{t}</Page>\n")


def таблиця(імя, шлях, t, колонки, ro=False, висота=None, vstretch=None, cmdbar=True, rowset=True, події_=None):
    """Table: CommandBarLocation → ReadOnly → ChangeRowSet → ChangeRowOrder → HeightInTableRows → VerticalStretch →
    DataPath → RowFilter → ContextMenu → AutoCommandBar → ExtendedTooltip → Search*Additions → Events → ChildItems."""
    s = f"{t}<Table name=\"{імя}\" id=\"{ід()}\">\n"
    if not cmdbar:
        s += f"{t}\t<CommandBarLocation>None</CommandBarLocation>\n"
    if ro:
        s += f"{t}\t<ReadOnly>true</ReadOnly>\n"
    if not rowset:
        s += f"{t}\t<ChangeRowSet>false</ChangeRowSet>\n{t}\t<ChangeRowOrder>false</ChangeRowOrder>\n"
    if висота:
        s += f"{t}\t<HeightInTableRows>{висота}</HeightInTableRows>\n"
    if vstretch is not None:
        s += f"{t}\t<VerticalStretch>{b(vstretch)}</VerticalStretch>\n"
    s += (f"{t}\t<DataPath>{шлях}</DataPath>\n{t}\t<RowFilter xsi:nil=\"true\"/>\n"
          f"{t}\t<ContextMenu name=\"{імя}КонтекстноеМеню\" id=\"{ід()}\"/>\n"
          f"{t}\t<AutoCommandBar name=\"{імя}КоманднаяПанель\" id=\"{ід()}\"/>\n"
          f"{t}\t<ExtendedTooltip name=\"{імя}РасширеннаяПодсказка\" id=\"{ід()}\"/>\n")
    for суф, тег, тип_ in (("СтрокаПоиска", "SearchStringAddition", "SearchStringRepresentation"),
                           ("СостояниеПросмотра", "ViewStatusAddition", "ViewStatusRepresentation"),
                           ("УправлениеПоиском", "SearchControlAddition", "SearchControl")):
        s += (f"{t}\t<{тег} name=\"{імя}{суф}\" id=\"{ід()}\">\n{t}\t\t<AdditionSource>\n{t}\t\t\t<Item>{імя}</Item>\n"
              f"{t}\t\t\t<Type>{тип_}</Type>\n{t}\t\t</AdditionSource>\n"
              f"{t}\t\t<ContextMenu name=\"{імя}{суф}КонтекстноеМеню\" id=\"{ід()}\"/>\n"
              f"{t}\t\t<ExtendedTooltip name=\"{імя}{суф}РасширеннаяПодсказка\" id=\"{ід()}\"/>\n{t}\t</{тег}>\n")
    s += події(події_, t + "\t") + f"{t}\t<ChildItems>\n" + колонки(t + "\t\t") + f"{t}\t</ChildItems>\n{t}</Table>\n"
    return s


def перемикач(імя, шлях, t, варіанти, події_=None):
    """RadioButtonField (Tumbler): DataPath → TitleLocation → RadioButtonType → ChoiceList."""
    s = (f"{t}<RadioButtonField name=\"{імя}\" id=\"{ід()}\">\n{t}\t<DataPath>{шлях}</DataPath>\n"
         f"{t}\t<TitleLocation>None</TitleLocation>\n{t}\t<RadioButtonType>Tumbler</RadioButtonType>\n{t}\t<ChoiceList>\n")
    for значення, представлення in варіанти:
        u = t + "\t\t"
        s += (f"{u}<xr:Item>\n{u}\t<xr:Presentation/>\n{u}\t<xr:CheckState>0</xr:CheckState>\n"
              f"{u}\t<xr:Value xsi:type=\"FormChoiceListDesTimeValue\">\n" + текст("Presentation", представлення, u + "\t\t")
              + f"{u}\t\t<Value xsi:type=\"xs:string\">{значення}</Value>\n{u}\t</xr:Value>\n{u}</xr:Item>\n")
    return s + f"{t}\t</ChoiceList>\n" + хвіст(імя, t + "\t") + події(події_, t + "\t") + f"{t}</RadioButtonField>\n"


def html_поле(імя, шлях, t):
    return (f"{t}<HTMLDocumentField name=\"{імя}\" id=\"{ід()}\">\n{t}\t<DataPath>{шлях}</DataPath>\n"
            f"{t}\t<TitleLocation>None</TitleLocation>\n" + хвіст(імя, t + "\t") + f"{t}</HTMLDocumentField>\n")


def кнопка_панелі(імя, команда, pic, t, default=False):
    """Button: Type → Representation → DefaultButton → CommandName → Picture → ExtendedTooltip."""
    s = (f"{t}<Button name=\"{імя}\" id=\"{ід()}\">\n{t}\t<Type>CommandBarButton</Type>\n"
         f"{t}\t<Representation>PictureAndText</Representation>\n")
    if default:
        s += f"{t}\t<DefaultButton>true</DefaultButton>\n"
    s += f"{t}\t<CommandName>Form.Command.{команда}</CommandName>\n" + картинка(pic, t + "\t")
    return s + f"{t}\t<ExtendedTooltip name=\"{імя}РасширеннаяПодсказка\" id=\"{ід()}\"/>\n{t}</Button>\n"


# ---------------------------------------------------------------- команди
КОМАНДИ = [
    # (ім'я, заголовок, підказка, картинка, за замовчуванням)
    ("ПрочитатьФайл", "1. Прочитати і зіставити", "Прочитати файл накладної і підібрати картки товарів", "Reread", False),
    ("ПодтвердитьСхожие", "Підтвердити схожі ≥ 85 %", "Поставити ✓ жовтим рядкам зі схожістю 85 % і вище", "CheckAll", False),
    ("ЗагрузитьВДокумент", "2. Завантажити в документ", "Записати рядки з ✓ у товари документа реалізації", "ExecuteTask", True),
    ("ВыгрузитьВExcel", "Excel з кодами 1С", "Зберегти накладну з кодами підтверджених карток (наступного разу — без перевірки)",
     "SaveFile", False),
    ("ЗабытьСоответствие", "Забути вибір рядка", "Програма більше не підставлятиме цю картку для такої назви", "InputFieldClear",
     False),
]


def командна_панель(t):
    s = f"{t}<AutoCommandBar name=\"ФормаКоманднаяПанель\" id=\"-1\">\n{t}\t<Autofill>false</Autofill>\n{t}\t<ChildItems>\n"
    for імя, _, _, pic, default in КОМАНДИ:
        s += кнопка_панелі("Форма" + імя, імя, pic, t + "\t\t", default)
    return s + f"{t}\t</ChildItems>\n{t}</AutoCommandBar>\n"


# ---------------------------------------------------------------- вміст
def шапка(u):
    return (поле("Документ", "Объект.Документ", u, title="Реалізація", width=38, amw=False)
            + поле("ИмяФайла", "ИмяФайла", u, title="Файл накладної", width=50, amw=False, hstretch=True, choice=True,
                   textedit=False, hint="Виберіть файл Excel (xls, xlsx)…",
                   події_=[("StartChoice", "ИмяФайлаНачалоВыбора")]))


def колонки_розпізнавання(u):
    return (поле("РаспознаваниеНомерСтроки", "Распознавание.НомерСтроки", u, title="№", ro=True, width=3)
            + поле("РаспознаваниеНаименование", "Распознавание.Наименование", u, title="Назва з накладної", ro=True, width=34)
            + поле("РаспознаваниеКоличество", "Распознавание.Количество", u, title="Кількість", ro=True, width=9)
            + поле("РаспознаваниеЕдиница", "Распознавание.Единица", u, title="Од.", ro=True, width=4)
            + поле("РаспознаваниеЦена", "Распознавание.Цена", u, title="Ціна без ПДВ", ro=True, width=10)
            + поле("РаспознаваниеСумма", "Распознавание.Сумма", u, title="Сума без ПДВ", ro=True, width=11)
            + прапорець("РаспознаваниеПодтверждено", "Распознавание.Подтверждено", u, title="✓", edit_mode="EnterOnInput",
                        tooltip="Рядок буде завантажено в документ",
                        події_=[("OnChange", "РаспознаваниеПодтвержденоПриИзменении")])
            + поле("РаспознаваниеНоменклатура", "Распознавание.Номенклатура", u, title="Номенклатура 1С", width=34,
                   події_=[("OnChange", "РаспознаваниеНоменклатураПриИзменении")])
            + поле("РаспознаваниеСходство", "Распознавание.Сходство", u, title="Схожість, %", ro=True, width=5)
            + поле("РаспознаваниеСпособ", "Распознавание.Способ", u, title="Як знайдено", ro=True, width=17)
            + поле("РаспознаваниеОстаток", "Распознавание.Остаток", u, title="Залишок", ro=True, width=9)
            + поле("РаспознаваниеЕдиницаКарточки", "Распознавание.ЕдиницаКарточки", u, title="Од. картки", ro=True, width=5)
            + поле("РаспознаваниеСкладОстатка", "Распознавание.СкладОстатка", u, title="Склад залишку", ro=True, width=18)
            + поле("РаспознаваниеПояснение", "Распознавание.Пояснение", u, title="Пояснення", ro=True, width=40))


def підсумок(імя, шлях, title, u, width=12):
    return поле(імя, шлях, u, title=title, ro=True, title_loc="Top", width=width, amw=False)


def підсумки(u):
    return (група("ГруппаИтогиФайл", u, lambda v:
                  підсумок("ФайлСтрок", "ФайлСтрок", "Рядків", v, 6)
                  + підсумок("ФайлКоличество", "ФайлКоличество", "Кількість", v)
                  + підсумок("ФайлСумма", "ФайлСумма", "Сума без ПДВ", v, 14)
                  + підсумок("ФайлРазом", "ФайлРазом", "«Разом» у накладній", v, 14)
                  + напис_поле("КонтрольИтога", "КонтрольИтога", v, hstretch=False),
                  title="Усього в накладній", repr_="NormalSeparation")
            + група("ГруппаИтогиГотово", u, lambda v:
                    підсумок("ГотовоСтрок", "ГотовоСтрок", "Рядків", v, 6)
                    + підсумок("ГотовоКоличество", "ГотовоКоличество", "Кількість", v)
                    + підсумок("ГотовоСумма", "ГотовоСумма", "Сума без ПДВ", v, 14),
                    title="Готово до завантаження ✓", repr_="NormalSeparation")
            + група("ГруппаИтогиОсталось", u, lambda v:
                    підсумок("ОсталосьСтрок", "ОсталосьСтрок", "Рядків", v, 6)
                    + підсумок("ОсталосьКоличество", "ОсталосьКоличество", "Кількість", v)
                    + підсумок("ОсталосьСумма", "ОсталосьСумма", "Сума без ПДВ", v, 14),
                    title="Не підтверджено", repr_="NormalSeparation"))


def сторінка_зіставлення(u):
    return (група("ГруппаФильтр", u, lambda v:
                  перемикач("Показать", "Показать", v, [("Все", "Усі"), ("Решение", "Потребують рішення"),
                                                        ("НеНайдено", "Не знайдено"), ("Подтверждено", "Підтверджені")],
                            події_=[("OnChange", "ПоказатьПриИзменении")])
                  + напис("ЛегендаПодтверждено", " Підтверджено ", v)
                  + напис("ЛегендаПохоже", " Схоже — перевірте ", v)
                  + напис("ЛегендаВнеОстатков", " Поза залишками ", v)
                  + напис("ЛегендаНеНайдено", " Не знайдено ", v)
                  + напис("ЛегендаСнято", " Знято ", v))
            + таблиця("Распознавание", "Распознавание", u, колонки_розпізнавання, cmdbar=False, rowset=False,
                      події_=[("OnActivateRow", "РаспознаваниеПриАктивизацииСтроки")])
            + група("ГруппаИтоги", u, підсумки)
            + напис("ПодсказкаВарианты", "Варіанти для виділеного рядка — двічі клацніть, щоб підставити картку:", u)
            + таблиця("Кандидаты", "Кандидаты", u, lambda v:
                      поле("КандидатыНоменклатура", "Кандидаты.Номенклатура", v, title="Номенклатура 1С", width=40)
                      + поле("КандидатыСходство", "Кандидаты.Сходство", v, title="Схожість, %", width=5)
                      + поле("КандидатыОстаток", "Кандидаты.Остаток", v, title="Залишок", width=9)
                      + поле("КандидатыСклад", "Кандидаты.Склад", v, title="Склад залишку", width=24)
                      + поле("КандидатыПояснение", "Кандидаты.Пояснение", v, title="Чому запропоновано", width=50),
                      ro=True, висота=5, vstretch=False, cmdbar=False, події_=[("Selection", "КандидатыВыбор")]))


def сторінка_налаштувань(u):
    return (група("ГруппаКолонки", u, lambda v:
                  підсумок_ред("ПерваяСтрока", "Объект.ПерваяСтрока", "Перший рядок даних", v)
                  + підсумок_ред("КолонкаНаименование", "Объект.КолонкаНаименование", "Назва товару", v)
                  + підсумок_ред("КолонкаКоличество", "Объект.КолонкаКоличество", "Кількість", v)
                  + підсумок_ред("КолонкаЕдиница", "Объект.КолонкаЕдиница", "Одиниця", v)
                  + підсумок_ред("КолонкаЦена", "Объект.КолонкаЦена", "Ціна без ПДВ", v)
                  + підсумок_ред("КолонкаСумма", "Объект.КолонкаСумма", "Сума без ПДВ", v)
                  + підсумок_ред("КолонкаКод", "Объект.КолонкаКод", "Код 1С (0 — немає)", v),
                  title="Номери колонок у файлі Excel (A = 1, B = 2, C = 3 …)", repr_="NormalSeparation")
            + група("ГруппаОпции", u, lambda v:
                    прапорець("ИскатьВнеОстатков", "Объект.ИскатьВнеОстатков", v,
                              title="Шукати в довіднику, якщо на складах не знайдено", title_loc="Right")
                    + прапорець("ОчиститьТовары", "Объект.ОчиститьТовары", v,
                                title="Очищати товари документа перед завантаженням (рекомендовано: повторне "
                                      "завантаження замінює, а не дублює)", title_loc="Right"),
                    гориз=False, title="Пошук і завантаження", repr_="NormalSeparation")
            + група("ГруппаСклады", u, lambda v:
                    таблиця("Склады", "Объект.Склады", v, lambda w:
                            поле("СкладыСклад", "Объект.Склады.Склад", w, title="Склади для пошуку за залишками"),
                            висота=4, vstretch=False)
                    + напис("ПодсказкаСклады", "Картки товарів шукаються серед залишків цих складів. Склад документа "
                                               "додається автоматично. Налаштування запам'ятовуються.", v)))


def підсумок_ред(імя, шлях, title, u):
    return поле(імя, шлях, u, title=title, title_loc="Top", width=9, amw=False)


def вміст_форми(t):
    return (група("ГруппаШапка", t, шапка)
            + напис_поле("НадписьРезюме", "НадписьРезюме", t)
            + f"{t}<Pages name=\"Страницы\" id=\"{ід()}\">\n{t}\t<PagesRepresentation>TabsOnTop</PagesRepresentation>\n"
            + f"{t}\t<ExtendedTooltip name=\"СтраницыРасширеннаяПодсказка\" id=\"{ід()}\"/>\n{t}\t<ChildItems>\n"
            + сторінка("СтраницаСопоставление", "Зіставлення", "ListViewMode", t + "\t\t", сторінка_зіставлення)
            + сторінка("СтраницаНастройки", "Налаштування", "ReportSettings", t + "\t\t", сторінка_налаштувань)
            + сторінка("СтраницаИнструкция", "Інструкція", "Information", t + "\t\t",
                       lambda u: html_поле("ПолеИнструкции", "ТекстИнструкции", u))
            + f"{t}\t</ChildItems>\n{t}</Pages>\n")


# ---------------------------------------------------------------- реквізити форми
def тип(вид, t):
    if вид == "bool":
        return f"{t}<v8:Type>xs:boolean</v8:Type>\n"
    if вид.startswith("num"):
        d, f = вид[3:].split(",")
        return (f"{t}<v8:Type>xs:decimal</v8:Type>\n{t}<v8:NumberQualifiers>\n{t}\t<v8:Digits>{d}</v8:Digits>\n"
                f"{t}\t<v8:FractionDigits>{f}</v8:FractionDigits>\n{t}\t<v8:AllowedSign>Any</v8:AllowedSign>\n"
                f"{t}</v8:NumberQualifiers>\n")
    if вид.startswith("str"):
        n = вид[3:]
        return (f"{t}<v8:Type>xs:string</v8:Type>\n{t}<v8:StringQualifiers>\n{t}\t<v8:Length>{n}</v8:Length>\n"
                f"{t}\t<v8:AllowedLength>Variable</v8:AllowedLength>\n{t}</v8:StringQualifiers>\n")
    return f"{t}<v8:Type>cfg:{вид}</v8:Type>\n"


def атрибут(імя, вид, t="\t\t", main=False, колонки=None, title=None):
    s = f"{t}<Attribute name=\"{імя}\" id=\"{ід()}\">\n"
    if title:
        s += текст("Title", title, t + "\t")
    s += f"{t}\t<Type>\n" + (f"{t}\t\t<v8:Type>v8:ValueTable</v8:Type>\n" if колонки else тип(вид, t + "\t\t")) + f"{t}\t</Type>\n"
    if main:
        s += f"{t}\t<MainAttribute>true</MainAttribute>\n"
    if колонки:
        s += f"{t}\t<Columns>\n"
        for кім, квид in колонки:
            s += (f"{t}\t\t<Column name=\"{кім}\" id=\"{ід()}\">\n{t}\t\t\t<Type>\n" + тип(квид, t + "\t\t\t\t")
                  + f"{t}\t\t\t</Type>\n{t}\t\t</Column>\n")
        s += f"{t}\t</Columns>\n"
    return s + f"{t}</Attribute>\n"


КОЛОНКИ_РАСП = [("НомерСтроки", "num5,0"), ("Наименование", "str300"), ("Количество", "num15,5"), ("Единица", "str25"),
                ("Цена", "num15,2"), ("Сумма", "num15,2"), ("Код", "str25"), ("Номенклатура", "CatalogRef.Номенклатура"),
                ("Подтверждено", "bool"), ("Способ", "str50"), ("Сходство", "num3,0"), ("Остаток", "num15,3"),
                ("ЕдиницаКарточки", "str25"), ("СкладОстатка", "str150"), ("Пояснение", "str300"),
                ("АвтоНоменклатура", "CatalogRef.Номенклатура"), ("АвтоСпособ", "str50"), ("АвтоСходство", "num3,0"),
                ("Группа", "str20")]  # Группа — лише у формі: стан рядка для перемикача-фільтра
КОЛОНКИ_ВАР = [("НомерСтроки", "num5,0"), ("Номенклатура", "CatalogRef.Номенклатура"), ("Сходство", "num3,0"),
               ("Остаток", "num15,3"), ("Склад", "str150"), ("Пояснение", "str300")]
ПІДСУМКИ = [("ФайлСтрок", "num6,0"), ("ФайлКоличество", "num15,3"), ("ФайлСумма", "num15,2"), ("ФайлРазом", "num15,2"),
            ("ГотовоСтрок", "num6,0"), ("ГотовоКоличество", "num15,3"), ("ГотовоСумма", "num15,2"),
            ("ОсталосьСтрок", "num6,0"), ("ОсталосьКоличество", "num15,3"), ("ОсталосьСумма", "num15,2")]


def головний():
    тіло = вміст_форми("\t\t")
    атрибути = (атрибут("Объект", f"ExternalDataProcessorObject.{ИМЯ}", main=True)
                + атрибут("ИмяФайла", "str0") + атрибут("АдресФайла", "str0")
                + атрибут("НадписьРезюме", "str0") + атрибут("КонтрольИтога", "str0")
                + атрибут("ТекстИнструкции", "str0", title="Інструкція")
                + атрибут("Показать", "str20")
                + "".join(атрибут(і, в) for і, в in ПІДСУМКИ)
                + атрибут("Распознавание", None, колонки=КОЛОНКИ_РАСП)
                + атрибут("Варианты", None, колонки=КОЛОНКИ_ВАР)
                + атрибут("Кандидаты", None, колонки=КОЛОНКИ_ВАР[1:]))
    команди = "".join(f"\t\t<Command name=\"{і}\" id=\"{ід()}\">\n" + текст("Title", заг, "\t\t\t")
                      + текст("ToolTip", підк, "\t\t\t") + f"\t\t\t<Action>{і}</Action>\n\t\t</Command>\n"
                      for і, заг, підк, _, _ in КОМАНДИ)
    return (f'﻿<?xml version="1.0" encoding="UTF-8"?>\n<Form {NS_FORM}>\n'
            + текст("Title", "Завантаження накладної з Excel у Реалізацію", "\t")
            + "\t<AutoTitle>false</AutoTitle>\n\t<VerticalScroll>useIfNecessary</VerticalScroll>\n"
            + командна_панель("\t")
            + "\t<Events>\n\t\t<Event name=\"OnCreateAtServer\">ПриСозданииНаСервере</Event>\n\t</Events>\n"
            + "\t<ChildItems>\n" + тіло + "\t</ChildItems>\n\t<Attributes>\n" + атрибути + "\t</Attributes>\n"
            + "\t<Commands>\n" + команди + "\t</Commands>\n</Form>\n")


форма = головний()
шлях = os.path.join(КАТАЛОГ, "Forms", "Форма", "Ext", "Form.xml")
with open(шлях, "w", encoding="utf-8", newline="\r\n") as f:
    f.write(форма)
print("OK Form.xml, останній id", ід.n)
