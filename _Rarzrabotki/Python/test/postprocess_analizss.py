# -*- coding: utf-8 -*-
# Пост-обробка Templates/МакетАнализСС/Ext/Template.xml після skd-compile:
#  1) інжект uk-заголовків dataset-полів (skd-compile їх не задає) за макетом замовника (Excel);
#  2) ru -> uk для заголовків обчислюваних полів (calculatedField / settingsVariant), §34.
# Запис UTF-8 БЕЗ BOM. Ідемпотентно на СВІЖОМУ виводі skd-compile (він не містить dataset-title).
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

TEMPLATE = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Создания комплектаций для бух учета\Templates\МакетАнализСС\Ext\Template.xml"

# dataset field -> uk заголовок колонки (порядок колонок керується selection у JSON).
# ОбщееНазвание / Номенклатура лишаються AUTO (рендеряться як caption груп).
TITLES = {
    "ОбщееНазвание":    "Загальна назва",
    "Единица":          "Одиниця виміру залишків",
    "ЕдиницаСС":        "Одиниця виміру СС",
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
    t = open(TEMPLATE, encoding='utf-8-sig').read()  # utf-8-sig прибирає BOM, який додає skd-compile

    injected = []
    for name, title in TITLES.items():
        anchor = '<field>' + name + '</field>'
        if anchor not in t:
            print(f"  УВАГА: поле {name} не знайдено у Template")
            continue
        # не інжектити двічі: якщо одразу після <field> вже є <title> — пропустити
        idx = t.index(anchor) + len(anchor)
        tail = t[idx:idx + 40]
        if '<title' in tail:
            continue
        t = t.replace(anchor, anchor + title_block(title), 1)
        injected.append(name)

    # ru -> uk для calc-полів / варіанту (§34)
    ru_count = t.count('<v8:lang>ru</v8:lang>')
    t = t.replace('<v8:lang>ru</v8:lang>', '<v8:lang>uk</v8:lang>')

    # UTF-8 без BOM
    with open(TEMPLATE, 'w', encoding='utf-8', newline='') as f:
        f.write(t)

    print(f"OK: інжектовано dataset-title = {len(injected)} ({', '.join(injected)}); ru->uk = {ru_count}")


if __name__ == '__main__':
    main()
