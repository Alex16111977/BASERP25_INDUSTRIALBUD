# -*- coding: utf-8 -*-
"""Точечный патч формы элемента СтруктураСебестоимости (BuhBud), 2026-09-27.

Добавляет (id с 300, существующие элементы не трогаются):
  - кнопку «Файли» (CommonCommand.ПрисоединенныеФайлы) в командную панель формы;
  - событие OnCreateAtServer -> ПриСозданииНаСервере (условное оформление кодом);
  - колонку «Додано вручну» в таблицу Комплектующие;
  - мини-таблицу «Файли підстави» (Объект.ФайлыОснования, 3 строки, только просмотр) над оригиналом «По модульно»;
  - страницу «Перерахунок» с таблицей Объект.КоэффициентыПересчета и тремя командами;
  - команды ЗаполнитьКоэффициенты, ПересчитатьКоличество, ЗаписатьПравилоВСправочник.
Канон компоновки — навык bas-form-layout (одна резиновая таблица на страницу, мини-таблица 3 строки).
Повторный запуск запрещён (assert на отсутствие новых элементов).

Запуск: python form_patch_ss_koef_fajly.py <путь к Form.xml>
"""
import sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

_ids = iter(range(300, 1000))


def nid():
    return next(_ids)


def title(ru, uk, tag='Title', ind=''):
    return (f'{ind}<{tag}>\n{ind}\t<v8:item>\n{ind}\t\t<v8:lang>ru</v8:lang>\n{ind}\t\t<v8:content>{ru}</v8:content>\n'
            f'{ind}\t</v8:item>\n{ind}\t<v8:item>\n{ind}\t\t<v8:lang>uk</v8:lang>\n{ind}\t\t<v8:content>{uk}</v8:content>\n'
            f'{ind}\t</v8:item>\n{ind}</{tag}>\n')


def column(table, field, ind, ro=False, ttl=None, footer=False, events=None, kind='InputField'):
    name = f'{table}{field}'
    s = f'{ind}<{kind} name="{name}" id="{nid()}">\n'
    s += f'{ind}\t<DataPath>Объект.{table}.{field}</DataPath>\n'
    if ro:
        s += f'{ind}\t<ReadOnly>true</ReadOnly>\n'
    if ttl:
        s += title(ttl[0], ttl[1], ind=ind + '\t')
    s += f'{ind}\t<EditMode>EnterOnInput</EditMode>\n'
    if kind == 'CheckBoxField':
        s += f'{ind}\t<CheckBoxType>Auto</CheckBoxType>\n'
    if footer:
        s += f'{ind}\t<FooterDataPath>Объект.{table}.Total{field}</FooterDataPath>\n'
    s += f'{ind}\t<ContextMenu name="{name}КонтекстноеМеню" id="{nid()}"/>\n'
    s += f'{ind}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{nid()}"/>\n'
    if events:
        s += f'{ind}\t<Events>\n'
        for ev, handler in events:
            s += f'{ind}\t\t<Event name="{ev}">{handler}</Event>\n'
        s += f'{ind}\t</Events>\n'
    s += f'{ind}</{kind}>\n'
    return s


def additions(table, ind):
    s = ''
    for tag, suffix, kind in (('SearchStringAddition', 'СтрокаПоиска', 'SearchStringRepresentation'),
                              ('ViewStatusAddition', 'СостояниеПросмотра', 'ViewStatusRepresentation'),
                              ('SearchControlAddition', 'УправлениеПоиском', 'SearchControl')):
        s += (f'{ind}<{tag} name="{table}{suffix}" id="{nid()}">\n{ind}\t<AdditionSource>\n{ind}\t\t<Item>{table}</Item>\n'
              f'{ind}\t\t<Type>{kind}</Type>\n{ind}\t</AdditionSource>\n'
              f'{ind}\t<ContextMenu name="{table}{suffix}КонтекстноеМеню" id="{nid()}"/>\n'
              f'{ind}\t<ExtendedTooltip name="{table}{suffix}РасширеннаяПодсказка" id="{nid()}"/>\n{ind}</{tag}>\n')
    return s


def line_number(table, ind):
    name = f'{table}НомерСтроки'
    return (f'{ind}<LabelField name="{name}" id="{nid()}">\n{ind}\t<DataPath>Объект.{table}.LineNumber</DataPath>\n'
            f'{ind}\t<EditMode>EnterOnInput</EditMode>\n{ind}\t<ContextMenu name="{name}КонтекстноеМеню" id="{nid()}"/>\n'
            f'{ind}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{nid()}"/>\n{ind}</LabelField>\n')


def button(name, command, ind):
    return (f'{ind}<Button name="{name}" id="{nid()}">\n{ind}\t<Type>CommandBarButton</Type>\n'
            f'{ind}\t<CommandName>{command}</CommandName>\n'
            f'{ind}\t<ExtendedTooltip name="{name}РасширеннаяПодсказка" id="{nid()}"/>\n{ind}</Button>\n')


def files_table(ind):
    t = 'ФайлыОснования'
    i2 = ind + '\t'
    s = f'{ind}<Table name="{t}" id="{nid()}">\n'
    s += f'{i2}<CommandBarLocation>None</CommandBarLocation>\n{i2}<ReadOnly>true</ReadOnly>\n'
    s += f'{i2}<HeightInTableRows>3</HeightInTableRows>\n{i2}<VerticalStretch>false</VerticalStretch>\n'
    s += f'{i2}<DataPath>Объект.{t}</DataPath>\n'
    s += title('Файлы основания (оригиналы Excel; модуль 0 — на всю СС; открыть — кнопка «Файлы»)',
               'Файли підстави (оригінали Excel; модуль 0 — на всю СС; відкрити — кнопка «Файли»)', ind=i2)
    s += f'{i2}<RowFilter xsi:nil="true"/>\n'
    s += f'{i2}<ContextMenu name="{t}КонтекстноеМеню" id="{nid()}"/>\n'
    s += f'{i2}<AutoCommandBar name="{t}КоманднаяПанель" id="{nid()}"/>\n'
    s += f'{i2}<ExtendedTooltip name="{t}РасширеннаяПодсказка" id="{nid()}"/>\n'
    s += additions(t, i2)
    s += f'{i2}<ChildItems>\n'
    i3 = i2 + '\t'
    for f in ('Модуль', 'Файл', 'ДатаЗагрузки', 'КоличествоСтрок', 'СуммаИтого'):
        s += column(t, f, i3)
    s += f'{i2}</ChildItems>\n{ind}</Table>\n'
    return s


def koef_page(ind):
    t = 'КоэффициентыПересчета'
    i1, i2, i3, i4 = ind + '\t', ind + '\t\t', ind + '\t\t\t', ind + '\t\t\t\t'
    s = f'{ind}<Page name="СтраницаПерерахунок" id="{nid()}">\n'
    s += title('Пересчет', 'Перерахунок', ind=i1)
    s += title('Коэффициенты пересчета единиц оригинала в единицы закупки',
               'Коефіцієнти перерахунку одиниць оригіналу в одиниці закупівлі', tag='ToolTip', ind=i1)
    s += f'{i1}<ExtendedTooltip name="СтраницаПерерахунокРасширеннаяПодсказка" id="{nid()}"/>\n'
    s += f'{i1}<ChildItems>\n'
    s += f'{i2}<Table name="{t}" id="{nid()}">\n'
    s += f'{i3}<Representation>List</Representation>\n{i3}<ChangeRowSet>false</ChangeRowSet>\n'
    s += f'{i3}<Footer>true</Footer>\n{i3}<DataPath>Объект.{t}</DataPath>\n'
    s += f'{i3}<RowFilter xsi:nil="true"/>\n'
    s += f'{i3}<ContextMenu name="{t}КонтекстноеМеню" id="{nid()}"/>\n'
    s += f'{i3}<AutoCommandBar name="{t}КоманднаяПанель" id="{nid()}">\n{i4}<ChildItems>\n'
    for name in ('ЗаполнитьКоэффициенты', 'ПересчитатьКоличество', 'ЗаписатьПравилоВСправочник'):
        s += button(f'{t}{name}', f'Form.Command.{name}', i4 + '\t')
    s += f'{i4}</ChildItems>\n{i3}</AutoCommandBar>\n'
    s += f'{i3}<ExtendedTooltip name="{t}РасширеннаяПодсказка" id="{nid()}"/>\n'
    s += additions(t, i3)
    s += f'{i3}<ChildItems>\n'
    s += line_number(t, i4)
    s += column(t, 'Этап', i4, ro=True)
    s += column(t, 'ОбщееНазвание', i4, ro=True)
    s += column(t, 'НоменклатураСС', i4, ro=True)
    s += column(t, 'КоличествоОригинал', i4, ro=True, ttl=('Кол-во оригинал', 'Кількість оригінал'), footer=True)
    s += column(t, 'ЕдиницаСС', i4, ro=True, ttl=('Ед. оригинала', 'Од. оригіналу'))
    s += column(t, 'Коэффициент', i4, events=[('OnChange', 'КоэффициентыПересчетаКоэффициентПриИзменении')])
    s += column(t, 'КоличествоПересчитанное', i4, ro=True, ttl=('Кол-во закупки', 'Кількість закупівлі'))
    s += column(t, 'ЕдиницаЗакупки', i4, ro=True, ttl=('Ед. закупки', 'Од. закупівлі'))
    s += column(t, 'ИсточникКоэффициента', i4, ro=True, ttl=('Источник', 'Джерело'))
    s += f'{i3}</ChildItems>\n{i2}</Table>\n{i1}</ChildItems>\n{ind}</Page>\n'
    return s


def command(name, ru, uk, tip_ru, tip_uk, ind='\t\t'):
    return (f'{ind}<Command name="{name}" id="{{CMDID}}">\n' + title(ru, uk, ind=ind + '\t')
            + title(tip_ru, tip_uk, tag='ToolTip', ind=ind + '\t') + f'{ind}\t<Action>{name}</Action>\n{ind}</Command>\n')


def replace_once(text, old, new):
    assert text.count(old) == 1, ('не найдено ровно 1 раз', old[:80])
    return text.replace(old, new)


def main(path):
    raw = open(path, 'rb').read()
    text = raw.decode('utf-8-sig')
    crlf = '\r\n' in text
    text = text.replace('\r\n', '\n')
    assert 'СтраницаПерерахунок' not in text and 'ФайлыОснования' not in text, 'патч уже применён'
    assert 'id="300"' not in text

    # 1. Кнопка «Файлы» в командной панели формы.
    text = replace_once(text, '\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1"/>\n',
                        '\t<AutoCommandBar name="ФормаКоманднаяПанель" id="-1">\n\t\t<ChildItems>\n'
                        f'\t\t\t<Button name="ФормаОбщаяКомандаПрисоединенныеФайлы" id="{nid()}">\n'
                        '\t\t\t\t<Type>CommandBarButton</Type>\n'
                        '\t\t\t\t<CommandName>CommonCommand.ПрисоединенныеФайлы</CommandName>\n'
                        '\t\t\t\t<DataPath>Объект.Ref</DataPath>\n'
                        f'\t\t\t\t<ExtendedTooltip name="ФормаОбщаяКомандаПрисоединенныеФайлыРасширеннаяПодсказка" id="{nid()}"/>\n'
                        '\t\t\t</Button>\n\t\t</ChildItems>\n\t</AutoCommandBar>\n')

    # 2. Событие ПриСозданииНаСервере.
    text = replace_once(text, '\t\t<Event name="BeforeWrite">ПередЗаписью</Event>\n',
                        '\t\t<Event name="OnCreateAtServer">ПриСозданииНаСервере</Event>\n'
                        '\t\t<Event name="BeforeWrite">ПередЗаписью</Event>\n')

    # 3. Колонка «Додано вручну» после «Не совпадает» в Комплектующие.
    anchor = '\t\t\t\t\t\t\t\t\t<ExtendedTooltip name="КомплектующиеНеСовпадаетРасширеннаяПодсказка" id="69"/>\n\t\t\t\t\t\t\t\t</CheckBoxField>\n'
    text = replace_once(text, anchor, anchor + column('Комплектующие', 'ДобавленоВручную', '\t' * 8, ro=True,
                                                      kind='CheckBoxField'))

    # 4. Мини-таблица «Файлы основания» перед оригиналом на странице «По модульно».
    anchor = '\t\t\t\t\t\t<Table name="КомплектующиеПоМодульно" id="170">\n'
    text = replace_once(text, anchor, files_table('\t' * 6) + anchor)

    # 5. Страница «Перерахунок» после страницы «По модульно» (перед «Додаткова»).
    anchor = '\t\t\t\t<Page name="СтраницаДоплнительная" id="165">\n'
    text = replace_once(text, anchor, koef_page('\t' * 4) + anchor)

    # 6. Команды формы.
    cmds = (command('ЗаполнитьКоэффициенты', 'Заполнить коэффициенты', 'Заповнити коефіцієнти',
                    'Позиции оригинала с несовпадающими единицами; коэффициенты — из правил общих названий (ручные не перезаписываются)',
                    'Позиції оригіналу з різними одиницями; коефіцієнти — з правил загальних назв (ручні не перезаписуються)')
            + command('ПересчитатьКоличество', 'Пересчитать количество', 'Перерахувати кількість',
                      'Количество Комплектующих = оригинал × коэффициент; сумма не меняется',
                      'Кількість Комплектуючих = оригінал × коефіцієнт; сума не змінюється')
            + command('ЗаписатьПравилоВСправочник', 'Записать в справочник', 'Записати в довідник',
                      'Сохранить коэффициент текущей строки как правило общего названия',
                      'Зберегти коефіцієнт поточного рядка як правило загальної назви'))
    for cid in (2, 3, 4):
        cmds = cmds.replace('{CMDID}', str(cid), 1)
    text = replace_once(text, '\t</Commands>\n</Form>', cmds + '\t</Commands>\n</Form>')

    if crlf:
        text = text.replace('\n', '\r\n')
    open(path, 'wb').write(b'\xef\xbb\xbf' + text.encode('utf-8'))
    print('OK, последний id =', next(_ids) - 1)


if __name__ == '__main__':
    main(sys.argv[1])
