# -*- coding: utf-8 -*-
"""SQL-маппинг: А_ВидыКонтрагентовДляБаланса + реквизиты ДоговорыКонтрагентов для snowflake."""
import json, sys, os
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

base = os.path.join(os.path.dirname(__file__), '..', 'mapping', 'baserp_storage.json')
objects = json.load(open(base, encoding='utf-8'))['objects']

vk = objects['Справочник.А_ВидыКонтрагентовДляБаланса']
print(f"А_ВидыКонтрагентовДляБаланса: {vk['primary_table']}")

dog = objects['Справочник.ДоговорыКонтрагентов']
print(f"ДоговорыКонтрагентов: {dog['primary_table']}")
for f, v in dog['fields'].items():
    if f in ('А_ВидКонтрагента', 'А_НаправлениеОказаниеУслуг', 'А_СтатьяБаланса',
             'А_ФинАгент', 'ТипДоговора', 'Ссылка'):
        print(f"    {f} -> {v}")
