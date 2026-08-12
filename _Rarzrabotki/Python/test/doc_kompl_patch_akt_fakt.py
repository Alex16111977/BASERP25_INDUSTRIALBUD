# -*- coding: utf-8 -*-
"""Патч форми: друга кнопка Акта («за проведеними») + перейменування першої.

Ідемпотентний. Точковий патч наявного Form.xml (генератори не перезапускати, §Д7).
"""
import io
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

ФАЙЛ = (r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh"
        r"\Documents\РасчетКомплектаций\Forms\ФормаДокумента\Ext\Form.xml")

_ид = [0]


def nid():
    _ид[0] += 1
    return _ид[0]


def назва(текст, о):
    return (f"{о}<Title>\n{о}\t<v8:item>\n{о}\t\t<v8:lang>uk</v8:lang>\n"
            f"{о}\t\t<v8:content>{текст}</v8:content>\n{о}\t</v8:item>\n{о}</Title>")


def main():
    s = io.open(ФАЙЛ, encoding='utf-8-sig').read()

    if 'name="КнопкаПечатьАктФакт"' in s:
        print("Кнопка «Акт за проведеними» вже є — нічого не роблю.")
        return

    _ид[0] = max(int(x) for x in re.findall(r'id="(\d+)"', s))
    print(f"  максимальний зайнятий id: {_ид[0]}")

    # 1. перейменувати наявну кнопку Акта
    старе = "Акт понаднормового списання"
    нове = "Акт понаднормового списання (план)"
    if старе in s and нове not in s:
        s = s.replace(f"<v8:content>{старе}</v8:content>",
                      f"<v8:content>{нове}</v8:content>")
        print(f"  перейменовано: «{старе}» -> «{нове}»")

    # 2. кнопка після наявної КнопкаПечатьАкт
    m = re.search(r'(\n(\t+)<Button name="КнопкаПечатьАкт" id="\d+">'
                  r'(?:(?!</Button>).)*?</Button>)', s, re.S)
    if m is None:
        raise SystemExit("НЕ ЗНАЙДЕНО КнопкаПечатьАкт")
    т = m.group(2)
    кнопка = (f'\n{т}<Button name="КнопкаПечатьАктФакт" id="{nid()}">\n'
              f"{т}\t<Type>UsualButton</Type>\n"
              f"{т}\t<CommandName>Form.Command.ПечатьАктФакт</CommandName>\n"
              f"{назва('Акт понаднормового списання (за проведеними)', т + chr(9))}\n"
              f'{т}\t<ExtendedTooltip name="КнопкаПечатьАктФактРасширеннаяПодсказка"'
              f' id="{nid()}"/>\n'
              f"{т}</Button>")
    s = s[:m.end(1)] + кнопка + s[m.end(1):]
    print("  + кнопка «Акт понаднормового списання (за проведеними)»")

    # 3. команда
    m = re.search(r'\n(\t+)</Commands>', s)
    команда = (f'\n\t\t<Command name="ПечатьАктФакт" id="{nid()}">\n'
               f"{назва('Акт понаднормового списання (за проведеними)', chr(9) * 3)}\n"
               f"\t\t\t<Action>ПечатьАктФакт</Action>\n\t\t</Command>")
    s = s[:m.start()] + команда + s[m.start():]
    print("  + команда ПечатьАктФакт")

    with io.open(ФАЙЛ, "w", encoding='utf-8-sig', newline="\n") as f:
        f.write(s)

    ids = re.findall(r'id="(-?\d+)"', s)
    дубли = [i for i in set(ids) if ids.count(i) > 1]
    print(f"\nЗаписано. Елементів: {len(ids)}, дублі id: {дубли if дубли else 'немає'}")


if __name__ == "__main__":
    main()
