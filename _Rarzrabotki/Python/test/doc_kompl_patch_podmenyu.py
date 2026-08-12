# -*- coding: utf-8 -*-
"""Патч форми Документ.РасчетКомплектаций: підменю «Документи» + колонки «Проведено».

Точковий патч наявного Form.xml (генератори форми НЕ перезапускати — §Д7).
Ідемпотентний: якщо елементи вже є, нічого не робить.

Додає:
  * Popup «Документи» у командну панель: Оновити документи / Розпровести всі /
    Провести всі / Оновити стан;
  * 4 Command + Action;
  * по 2 CheckBoxField у таблиці ДокументиКомплектації та ДокументиМалоценки.

Спека: docs/superpowers/specs/2026-08-04-podmenyu-dokumenty-design.md
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


def кнопка(имя, команда, титул, о):
    т = "\t" * о
    return (f'{т}<Button name="{имя}" id="{nid()}">\n'
            f"{т}\t<Type>UsualButton</Type>\n"
            f"{т}\t<CommandName>Form.Command.{команда}</CommandName>\n"
            f"{назва(титул, т + chr(9))}\n"
            f'{т}\t<ExtendedTooltip name="{имя}РасширеннаяПодсказка" id="{nid()}"/>\n'
            f"{т}</Button>")


def галочка(имя, путь, титул, о):
    т = "\t" * о
    return (f'{т}<CheckBoxField name="{имя}" id="{nid()}">\n'
            f"{т}\t<DataPath>{путь}</DataPath>\n"
            f"{т}\t<ReadOnly>true</ReadOnly>\n"
            f"{назва(титул, т + chr(9))}\n"
            f"{т}\t<CheckBoxType>Auto</CheckBoxType>\n"
            f'{т}\t<ContextMenu name="{имя}КонтекстноеМеню" id="{nid()}"/>\n'
            f'{т}\t<ExtendedTooltip name="{имя}РасширеннаяПодсказка" id="{nid()}"/>\n'
            f"{т}</CheckBoxField>")


КОМАНДЫ = [
    ("ОбновитьДокументы", "Оновити документи"),
    ("РаспровестиВсе", "Розпровести всі"),
    ("ПровестиВсе", "Провести всі"),
    ("ОбновитьСостояние", "Оновити стан"),
]


def main():
    s = io.open(ФАЙЛ, encoding='utf-8-sig').read()
    исходный = s

    if 'name="ПодменюДокументы"' in s:
        print("Підменю «Документи» вже є — нічого не роблю.")
        return

    # стартуємо нумерацію з max+1
    _ид[0] = max(int(x) for x in re.findall(r'id="(\d+)"', s))
    print(f"  максимальний зайнятий id: {_ид[0]}")

    # ---- 1. Popup у командну панель: перед кнопкою ФормаСправка ----
    попап_ид = nid()
    попап_подсказка = nid()
    кнопки = "\n".join(кнопка("Кнопка" + и, и, т, 6) for и, т in КОМАНДЫ)
    попап = (f'\t\t\t\t<Popup name="ПодменюДокументы" id="{попап_ид}">\n'
             f"{назва('Документи', chr(9) * 5)}\n"
             f"\t\t\t\t\t<Representation>Auto</Representation>\n"
             f'\t\t\t\t\t<ExtendedTooltip name="ПодменюДокументыРасширеннаяПодсказка"'
             f' id="{попап_подсказка}"/>\n'
             f"\t\t\t\t\t<ChildItems>\n{кнопки}\n\t\t\t\t\t</ChildItems>\n"
             f"\t\t\t\t</Popup>\n")

    m = re.search(r'\n(\t+)<Button name="ФормаСправка"', s)
    if m is None:
        raise SystemExit("НЕ ЗНАЙДЕНО кнопку ФормаСправка — куди вставляти Popup?")
    s = s[:m.start() + 1] + попап + s[m.start() + 1:]
    print("  + Popup «Документи» перед кнопкою ФормаСправка")

    # ---- 2. Колонки-галочки в таблицях документів ----
    for таблица, префикс in (("ДокументиКомплектації", "Компл"),
                             ("ДокументиМалоценки", "Мал")):
        шаблон = (r'(<Table name="' + re.escape(таблица) + r'" id="\d+">'
                  r'(?:(?!</Table>).)*?)(\n(\t+)</ChildItems>\n\t+</Table>)')
        m = re.search(шаблон, s, re.S)
        if m is None:
            print(f"  ⚠ таблиця {таблица} не знайдена — пропуск")
            continue
        отступ = len(m.group(3)) + 1
        поля = "\n".join([
            галочка(f"{таблица}ПроведенПоНормам",
                    f"Объект.{таблица}.ПроведенПоНормам", "Проведений (норма)", отступ),
            галочка(f"{таблица}ПроведенДодаткова",
                    f"Объект.{таблица}.ПроведенДодаткова", "Проведений (понад)", отступ),
        ])
        s = s[:m.end(1)] + "\n" + поля + s[m.end(1):]
        print(f"  + 2 колонки «Проведено» у таблицю {таблица}")

    # ---- 3. Commands ----
    команды_xml = "\n".join(
        f'\t\t<Command name="{и}" id="{nid()}">\n'
        f"{назва(т, chr(9) * 3)}\n"
        f"\t\t\t<Action>{и}</Action>\n\t\t</Command>"
        for и, т in КОМАНДЫ)
    m = re.search(r'\n(\t+)</Commands>', s)
    if m is None:
        raise SystemExit("НЕ ЗНАЙДЕНО </Commands>")
    s = s[:m.start()] + "\n" + команды_xml + s[m.start():]
    print(f"  + {len(КОМАНДЫ)} команди")

    if s == исходный:
        print("Змін немає.")
        return

    with io.open(ФАЙЛ, "w", encoding='utf-8-sig', newline="\n") as f:
        f.write(s)

    ids = re.findall(r'id="(-?\d+)"', s)
    дубли = [i for i in set(ids) if ids.count(i) > 1]
    print(f"\nЗаписано: {ФАЙЛ}")
    print(f"  елементів: {len(ids)}, дублі id: {дубли if дубли else 'немає'}")
    print(f"  максимальний id тепер: {_ид[0]}")


if __name__ == "__main__":
    main()
