# -*- coding: utf-8 -*-
# Пост-обробка Templates/МакетАнализССОдна/Ext/Template.xml після skd-compile:
#  1) інжект uk-заголовків dataset-полів (одна колонка одиниці «Одиниця виміру»);
#  2) ru -> uk для заголовків обчислюваних полів. UTF-8 БЕЗ BOM (utf-8-sig читання).
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

TEMPLATE = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Создания комплектаций для бух учета\Templates\МакетАнализССОдна\Ext\Template.xml"

TITLES = {
    "ОбщееНазвание":    "Загальна назва",
    "ЕдиницаЕдина":     "Одиниця виміру",
    "Остаток":          "Залишок",
    "НормаНоменкл":     "Згідно з СС",
    "ВНорме":           "В нормі",
    "ПонадНорму":       "Понад норму",
    "Экономия":         "Економія",
    "ВНормеСумма":      "В нормі, грн",
    "ПонадНормуСумма":  "Понад норму, грн",
    "ЕкономіяСума":     "Економія, грн",
    "НормаСуммаНоменкл": "Сума СС",
    "СуммаОстатка":     "Факт. списання",
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


def main():
    t = open(TEMPLATE, encoding='utf-8-sig').read()
    injected = []
    for name, title in TITLES.items():
        anchor = '<field>' + name + '</field>'
        if anchor not in t:
            print(f"  УВАГА: поле {name} не знайдено у Template")
            continue
        idx = t.index(anchor) + len(anchor)
        if '<title' in t[idx:idx + 40]:
            continue
        t = t.replace(anchor, anchor + title_block(title), 1)
        injected.append(name)
    ru = t.count('<v8:lang>ru</v8:lang>')
    t = t.replace('<v8:lang>ru</v8:lang>', '<v8:lang>uk</v8:lang>')
    with open(TEMPLATE, 'w', encoding='utf-8', newline='') as f:
        f.write(t)
    print(f"OK: dataset-title = {len(injected)} ({', '.join(injected)}); ru->uk = {ru}")


if __name__ == '__main__':
    main()
