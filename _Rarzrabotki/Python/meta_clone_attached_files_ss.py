# -*- coding: utf-8 -*-
"""Клон типового справочника присоединённых файлов под СтруктураСебестоимости (BuhBud), 2026-09-27.

Эталон: Catalogs/ДоговорыКонтрагентовПрисоединенныеФайлы (БСП «Работа с файлами»).
Результат: Catalogs/СтруктураСебестоимостиПрисоединенныеФайлы — имя по правилу БСП
«<Владелец>ПрисоединенныеФайлы», тип ВладелецФайла = CatalogRef.СтруктураСебестоимости,
все uuid (объект, атрибуты, форма, TypeId/ValueId) — новые.

Запуск: python meta_clone_attached_files_ss.py <каталог BASEBuh>
"""
import os
import re
import shutil
import sys
import uuid

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

SRC = 'ДоговорыКонтрагентовПрисоединенныеФайлы'
DST = 'СтруктураСебестоимостиПрисоединенныеФайлы'
OWNER_SRC = 'cfg:CatalogRef.ДоговорыКонтрагентов<'
OWNER_DST = 'cfg:CatalogRef.СтруктураСебестоимости<'
SYN = {
    'Присоединенные файлы (Договоры контрагентов)': 'Присоединенные файлы (Структура себестоимости)',
    'Приєднані файли (Договори контрагентів)': 'Приєднані файли (Структура собівартості)',
    'Присоединенные файлы (Договор контрагента)': 'Присоединенный файл (Структура себестоимости)',
    'Приєднані файли (Договір контрагента)': 'Приєднаний файл (Структура собівартості)',
}
HELP = {
    'ru.html': ('Справочник "Присоединенные файлы (Структура себестоимости)"',
                'Хранит исходные файлы (Excel), на основании которых введена структура себестоимости: '
                'на всю структуру или на отдельный модуль (см. табличную часть «Файлы основания»). '
                'Справочник служебный.'),
    'uk.html': ('Довідник "Приєднані файли (Структура собівартості)"',
                'Зберігає вихідні файли (Excel), на підставі яких введено структуру собівартості: '
                'на всю структуру або на окремий модуль (див. табличну частину «Файли підстави»). '
                'Довідник службовий.'),
}
HELP_TEMPLATE = ('<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"><html><head>'
                 '<meta content="text/html;charset=utf-8" http-equiv="content-type"></meta>'
                 '<link rel="stylesheet" type="text/css" href="v8help://service_book/service_style"></link>'
                 '</head><body>\n<h1>{0}</h1>\n<p>{1}</p></body></html>')
UUID_RE = re.compile(r'[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}')


def main(base):
    cat = os.path.join(base, 'Catalogs')
    src_xml, dst_xml = os.path.join(cat, SRC + '.xml'), os.path.join(cat, DST + '.xml')
    src_dir, dst_dir = os.path.join(cat, SRC), os.path.join(cat, DST)
    assert not os.path.exists(dst_xml) and not os.path.exists(dst_dir), 'приёмник уже существует'

    mapping = {}

    def new_uuid(m):
        old = m.group(0)
        if old not in mapping:
            mapping[old] = str(uuid.uuid4())
        return mapping[old]

    def transform(text, is_root):
        assert SRC in text or not is_root
        text = text.replace(SRC, DST)
        if is_root:
            assert text.count(OWNER_SRC) == 1, 'тип ВладелецФайла не найден ровно 1 раз'
            for a, b in SYN.items():
                assert a in text, 'синоним не найден: ' + a
                text = text.replace(a, b)
        text = text.replace(OWNER_SRC, OWNER_DST)
        return UUID_RE.sub(new_uuid, text)

    def copy_file(src, dst, is_root=False):
        with open(src, 'rb') as f:
            raw = f.read()
        bom = raw.startswith(b'\xef\xbb\xbf')
        text = raw[3:].decode('utf-8') if bom else raw.decode('utf-8')
        if src.endswith('.html') and os.path.basename(src) in HELP:
            text = HELP_TEMPLATE.format(*HELP[os.path.basename(src)])
        elif src.endswith(('.xml', '.bsl')):
            text = transform(text, is_root)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        with open(dst, 'wb') as f:
            f.write((b'\xef\xbb\xbf' if bom else b'') + text.encode('utf-8'))

    copy_file(src_xml, dst_xml, is_root=True)
    for root, _, files in os.walk(src_dir):
        for name in files:
            s = os.path.join(root, name)
            copy_file(s, os.path.join(dst_dir, os.path.relpath(s, src_dir)))

    # Контроль: в клоне не осталось ни имени эталона, ни одного старого uuid.
    leftovers = []
    for p in [dst_xml] + [os.path.join(r, n) for r, _, fs in os.walk(dst_dir) for n in fs]:
        with open(p, encoding='utf-8-sig') as f:
            t = f.read()
        if SRC in t or 'Договор' in t:
            leftovers.append(p)
        if any(old in t for old in mapping):
            leftovers.append(p + ' (старый uuid)')
    assert not leftovers, leftovers
    print('OK: uuid заменено', len(mapping), '; файлы:', dst_xml, dst_dir)


if __name__ == '__main__':
    main(sys.argv[1])
