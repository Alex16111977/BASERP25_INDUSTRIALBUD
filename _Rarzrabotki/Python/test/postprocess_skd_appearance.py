# -*- coding: utf-8 -*-
# Звужує conditionalAppearance у 3 СКД-макетах: умовне оформлення (зелений/янтар) має діяти
# ЛИШЕ на дані (групи), а НЕ на шапку-заголовок / заголовки полів / параметри.
# Пряме присвоєння ЦветФона на вже виведеному ТабДок НЕ перебиває оформлення СКД (перевірено),
# тому єдиний надійний шлях — області застосування (useInHeader/useInFieldsHeader/useInParameters = DontUse).
# Прибираємо також <dcsset:presentation> (опційне; спрощує порядок схеми — як у типових звітах).
# UTF-8 БЕЗ BOM. Ідемпотентно.
import re, sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

BASE = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\Создания комплектаций для бух учета\Templates"
TEMPLATES = [
    BASE + r"\МакетАнализСС\Ext\Template.xml",
    BASE + r"\МакетАнализССОдна\Ext\Template.xml",
    BASE + r"\МакетПланФакт\Ext\Template.xml",
    BASE + r"\МакетПланФактЕтапи\Ext\Template.xml",
]

FLAGS = ("useInFieldsHeader", "useInHeader", "useInParameters")  # НЕ фарбувати шапку/заголовки полів/параметри


def scope(block):
    # 1) прибрати presentation (опційний елемент) — уникнути питання порядку схеми
    block = re.sub(r'[ \t]*<dcsset:presentation\b.*?</dcsset:presentation>\n?', '', block, flags=re.S)

    # 2) після кожного </dcsset:appearance> додати useIn*-DontUse (група/деталі лишаються — Auto)
    def add(m):
        indent = m.group(1)
        flags = "".join(indent + f"<dcsset:{n}>DontUse</dcsset:{n}>" for n in FLAGS)
        return m.group(0) + flags
    block = re.sub(r'(\n[ \t]*)</dcsset:appearance>', add, block)
    return block


def main():
    for path in TEMPLATES:
        t = open(path, encoding='utf-8-sig').read()
        m = re.search(r'<dcsset:conditionalAppearance>.*?</dcsset:conditionalAppearance>', t, re.S)
        name = path.split("\\")[-3]
        if not m:
            print(f"  {name}: conditionalAppearance не знайдено — пропуск")
            continue
        blk = m.group(0)
        if "useInHeader" in blk:
            print(f"  {name}: вже звужено — пропуск")
            continue
        new_blk = scope(blk)
        t = t[:m.start()] + new_blk + t[m.end():]
        with open(path, 'w', encoding='utf-8', newline='') as f:
            f.write(t)
        print(f"OK {name}: useIn*-DontUse додано, presentation прибрано")


if __name__ == '__main__':
    main()
