# -*- coding: utf-8 -*-
"""
Установка скила 1c-bas-erp-developer в Claude Desktop
"""
import os, shutil

src_skill = r'C:\Razrabotki\ИндастриалБуд\скилс\files\SKILL.md'
base      = r'C:\Users\support1c\.claude\plugins\marketplaces\claude-plugins-official\external_plugins\1c-bas-erp-developer'
plugin_dir = os.path.join(base, '.claude-plugin')
skill_dir  = os.path.join(base, 'skills', '1c-bas-erp-developer')

os.makedirs(plugin_dir, exist_ok=True)
os.makedirs(skill_dir, exist_ok=True)
print('[OK] Папки созданы')

plugin_json = """{
  "name": "1c-bas-erp-developer",
  "description": "Senior 1C BAS ERP 2.5 developer for INDUSTRIALBUD",
  "version": "1.0.0",
  "author": { "name": "Alex", "url": "https://github.com/Alex16111977" },
  "keywords": ["1c", "bsl", "bas-erp", "ukraine"]
}"""
with open(os.path.join(plugin_dir, 'plugin.json'), 'w', encoding='utf-8') as f:
    f.write(plugin_json)
print('[OK] plugin.json записан')

shutil.copy2(src_skill, os.path.join(skill_dir, 'SKILL.md'))
size = os.path.getsize(os.path.join(skill_dir, 'SKILL.md'))
print(f'[OK] SKILL.md скопирован ({size} bytes)')

print()
print('=== Структура ===')
for root, dirs, files in os.walk(base):
    lvl = root.replace(base, '').count(os.sep)
    print('  ' * lvl + os.path.basename(root) + '/')
    for fn in files:
        sz = os.path.getsize(os.path.join(root, fn))
        print('  ' * (lvl+1) + fn + f'  ({sz} bytes)')

print()
print('Готово! Перезапусти Claude Desktop.')
