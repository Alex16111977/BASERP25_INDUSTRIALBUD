# -*- coding: utf-8 -*-
"""Идемпотентный патчер украинских подписей.

meta-compile и form-compile жёстко пишут только <v8:lang>ru</v8:lang>, а в конфигурации
DefaultLanguage = Language.Украинский — поэтому без uk-перевода 1С показывает ИМЯ элемента
(кнопка «Рассчитать», пустые закладки, колонки «ДокументПроведен»).

Скрипт проходит по XML документа, перечислений, формы и справочника договоров и добавляет
пару <v8:item><v8:lang>uk</v8:lang><v8:content>…</v8:content></v8:item> в каждый
<Synonym>/<Title>, где uk ещё нет. Русские подписи остаются — их видит разработчик.

Запускать ПОСЛЕ любой перегенерации meta-compile/form-compile.
"""
import io
import json
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

БАЗА = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh"
СЛОВАРЬ = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\doc_finrep_uk.json"

ФАЙЛЫ = [
    r"Documents\А_ФинансовыйОтчетПроизводства.xml",
    r"Documents\А_ФинансовыйОтчетПроизводства\Forms\ФормаДокумента.xml",
    r"Documents\А_ФинансовыйОтчетПроизводства\Forms\ФормаДокумента\Ext\Form.xml",
    r"InformationRegisters\А_ПараметрыДоговоровФинотчета.xml",
    r"Catalogs\ДоговорыКонтрагентов.xml",
]

# в справочнике договоров трогаем ТОЛЬКО наши реквизиты — типовые подписи не переписываем
ТОЛЬКО_СВОИ = {r"Catalogs\ДоговорыКонтрагентов.xml": "А_"}


def загрузить_словарь():
    д = json.load(io.open(СЛОВАРЬ, encoding='utf-8'))
    плоский = {}
    for раздел, пары in д.items():
        if раздел.startswith("_") or not isinstance(пары, dict):
            continue
        for ключ, значение in пары.items():
            # первое вхождение выигрывает: специфичные разделы идут раньше общих
            плоский.setdefault(ключ, значение)
    return плоский


def экранировать(текст):
    return (текст.replace("&", "&amp;").replace("<", "&lt;")
                 .replace(">", "&gt;").replace('"', "&quot;"))


# Метаданные: <Name>Имя</Name> … <Synonym|Title> … </Synonym|Title>
ШАБЛОН = re.compile(
    r'(?P<голова><Name>(?P<имя>[^<]+)</Name>\s*'
    r'<(?P<тег>Synonym|Title)>)(?P<тело>.*?)(?P<закрытие></(?P=тег)>)',
    re.S)

# Форма: имя лежит в АТРИБУТЕ, а между ним и <Title> есть промежуточные теги
# (<Type>, <CommandName>…). Негативный просмотр не даёт «убежать» в следующий элемент.
ШАБЛОН_ФОРМЫ = re.compile(
    r'(?P<голова><\w+ name="(?P<имя>[^"]+)"[^>]*>'
    r'(?:(?!<\w+ name=)[\s\S])*?'
    r'<(?P<тег>Title)>)(?P<тело>.*?)(?P<закрытие></(?P=тег)>)',
    re.S)


def патчить_файл(путь, словарь, префикс_фильтра):
    полный = БАЗА + "\\" + путь
    try:
        текст = io.open(полный, encoding='utf-8-sig').read()
    except FileNotFoundError:
        return None
    добавлено, пропущено, нет_перевода = 0, 0, []

    def замена(m):
        nonlocal добавлено, пропущено
        имя, тело = m.group("имя"), m.group("тело")
        if префикс_фильтра and not имя.startswith(префикс_фильтра):
            return m.group(0)
        if "<v8:lang>uk</v8:lang>" in тело:
            пропущено += 1
            return m.group(0)
        перевод = словарь.get(имя)
        if перевод is None:
            нет_перевода.append(имя)
            return m.group(0)
        отступ = "\t\t\t\t\t"
        вставка = (f"{тело.rstrip()}\n{отступ}<v8:item>\n{отступ}\t<v8:lang>uk</v8:lang>\n"
                   f"{отступ}\t<v8:content>{экранировать(перевод)}</v8:content>\n"
                   f"{отступ}</v8:item>\n{отступ[:-1]}")
        добавлено += 1
        return m.group("голова") + вставка + m.group("закрытие")

    новый = ШАБЛОН.sub(замена, текст)
    if путь.endswith("Form.xml"):
        новый = ШАБЛОН_ФОРМЫ.sub(замена, новый)
    if новый != текст:
        io.open(полный, 'w', encoding='utf-8-sig', newline='').write(новый)
    return добавлено, пропущено, sorted(set(нет_перевода))


def счётчики(путь):
    полный = БАЗА + "\\" + путь
    try:
        т = io.open(полный, encoding='utf-8-sig').read()
    except FileNotFoundError:
        return 0, 0
    return t_count(т, "ru"), t_count(т, "uk")


def t_count(текст, язык):
    return текст.count(f"<v8:lang>{язык}</v8:lang>")


def main():
    словарь = загрузить_словарь()
    print(f"Словник: {len(словарь)} записів\n")
    всего_добавлено = 0
    все_без_перевода = []

    for путь in ФАЙЛЫ:
        рез = патчить_файл(путь, словарь, ТОЛЬКО_СВОИ.get(путь))
        if рез is None:
            print(f"  — {путь.split(chr(92))[-1]:38} файл відсутній")
            continue
        добавлено, пропущено, без_перевода = рез
        ru, uk = счётчики(путь)
        всего_добавлено += добавлено
        имя_файла = путь.split(chr(92))[-1]
        родитель = путь.split(chr(92))[-2] if "\\" in путь else ""
        метка = f"{родитель}/{имя_файла}" if родитель in ("Ext", "ФормаДокумента") else имя_файла
        print(f"  + {метка:38} додано {добавлено:>3}, вже було {пропущено:>3}  |  ru={ru:>3} uk={uk:>3}")
        if без_перевода:
            все_без_перевода.extend(без_перевода)

    if все_без_перевода:
        уникальные = sorted(set(все_без_перевода))
        print(f"\n  ⚠ немає перекладу для {len(уникальные)} імен — лишились російськими:")
        for имя in уникальные:
            print(f"      {имя}")
        print("      додайте їх у doc_finrep_uk.json і запустіть скрипт ще раз")

    print(f"\nДодано uk-підписів: {всего_добавлено}")
    return 1 if все_без_перевода else 0


if __name__ == "__main__":
    sys.exit(main())
