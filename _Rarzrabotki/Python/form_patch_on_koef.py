# -*- coding: utf-8 -*-
"""Точечный патч формы элемента ОбщиеНазванияНоменклатуры (BuhBud), 2026-09-27:
мини-таблица «Коефіцієнти перерахунку» (Объект.КоэффициентыПересчета, 4 строки, редактируемая) перед таблицей
«Номенклатура». Правило: ЕдиницаСС × Коэффициент = Единица карточки; пустая НоменклатураСС — на всё название.
id с 100. Запуск: python form_patch_on_koef.py <путь к Form.xml>"""
import sys
sys.path.insert(0, __import__('os').path.dirname(__file__))
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
import form_patch_ss_koef_fajly as P

P._ids = iter(range(100, 300))


def main(path):
    raw = open(path, 'rb').read()
    text = raw.decode('utf-8-sig')
    crlf = '\r\n' in text
    text = text.replace('\r\n', '\n')
    assert 'КоэффициентыПересчета' not in text and 'id="100"' not in text, 'патч уже применён'
    t, ind = 'КоэффициентыПересчета', '\t\t'
    i2, i3 = ind + '\t', ind + '\t\t'
    s = f'{ind}<Table name="{t}" id="{P.nid()}">\n'
    s += f'{i2}<HeightInTableRows>4</HeightInTableRows>\n{i2}<VerticalStretch>false</VerticalStretch>\n'
    s += f'{i2}<DataPath>Объект.{t}</DataPath>\n'
    s += P.title('Коэффициенты пересчета: единица СС × коэффициент = единица общего названия (пустая номенклатура СС — на всё название)',
                 'Коефіцієнти перерахунку: одиниця СС × коефіцієнт = одиниця загальної назви (порожня номенклатура СС — на всю назву)', ind=i2)
    s += f'{i2}<RowFilter xsi:nil="true"/>\n'
    s += f'{i2}<ContextMenu name="{t}КонтекстноеМеню" id="{P.nid()}"/>\n'
    s += f'{i2}<AutoCommandBar name="{t}КоманднаяПанель" id="{P.nid()}"/>\n'
    s += f'{i2}<ExtendedTooltip name="{t}РасширеннаяПодсказка" id="{P.nid()}"/>\n'
    s += P.additions(t, i2)
    s += f'{i2}<ChildItems>\n'
    s += P.column(t, 'ЕдиницаСС', i3, ttl=('Единица СС', 'Одиниця СС'))
    s += P.column(t, 'НоменклатураСС', i3, ttl=('Номенклатура СС (пусто — всё название)', 'Номенклатура СС (порожньо — вся назва)'))
    s += P.column(t, 'Коэффициент', i3)
    s += P.column(t, 'Комментарий', i3)
    s += f'{i2}</ChildItems>\n{ind}</Table>\n'
    anchor = '\t\t<Table name="Номенклатура" id="7">\n'
    text = P.replace_once(text, anchor, s + anchor)
    if crlf:
        text = text.replace('\n', '\r\n')
    open(path, 'wb').write(b'\xef\xbb\xbf' + text.encode('utf-8'))
    print('OK')


if __name__ == '__main__':
    main(sys.argv[1])
