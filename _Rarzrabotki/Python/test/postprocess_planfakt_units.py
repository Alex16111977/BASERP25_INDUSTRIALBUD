# -*- coding: utf-8 -*-
# Пост-обробка Templates/{МакетПланФакт,МакетПланФактЕтапи}/Ext/Template.xml після skd-compile:
#  1) інжект uk-заголовків ЛИШЕ для 2 колонок одиниць (Единица / ЕдиницаСС); решта dataset-полів
#     лишаються авто-рос (як історично, зрозумілі);
#  2) ru -> uk для заголовків обчислюваних %-полів (§34).
# UTF-8 БЕЗ BOM (utf-8-sig читання прибирає BOM від skd-compile).
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

BASE = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Создания комплектаций для бух учета\Templates"
TEMPLATES = [BASE + r"\МакетПланФакт\Ext\Template.xml"]  # МакетПланФактЕтапи перенесено на postprocess_analizss_odna_etapy.py (Фаза 15)

TITLES = {
    "ОбщееНазвание": "Загальна назва",
    "Единица":   "Одиниця виміру залишків",
    "ЕдиницаСС": "Одиниця виміру СС",
}


def title_block(text):
    return (
        '\n\t\t\t<title xsi:type="v8:LocalStringType">'
        '\n\t\t\t\t<v8:item>'
        '\n\t\t\t\t\t<v8:lang>uk</v8:lang>'
        '\n\t\t\t\t\t<v8:content>' + text + '</v8:content>'
        '\n\t\t\t\t</v8:item>'
        '\n\t\t\t</title>'
    )


def process(path):
    t = open(path, encoding='utf-8-sig').read()
    injected = []
    for name, title in TITLES.items():
        anchor = '<field>' + name + '</field>'
        if anchor not in t:
            print(f"  УВАГА: {name} не знайдено у {path}")
            continue
        idx = t.index(anchor) + len(anchor)
        if '<title' in t[idx:idx + 40]:
            continue
        t = t.replace(anchor, anchor + title_block(title), 1)
        injected.append(name)
    ru = t.count('<v8:lang>ru</v8:lang>')
    t = t.replace('<v8:lang>ru</v8:lang>', '<v8:lang>uk</v8:lang>')
    with open(path, 'w', encoding='utf-8', newline='') as f:
        f.write(t)
    print(f"OK {path.split(chr(92))[-3]}: units={injected}; ru->uk={ru}")


if __name__ == '__main__':
    for p in TEMPLATES:
        process(p)
