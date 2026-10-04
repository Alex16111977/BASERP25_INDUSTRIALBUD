# -*- coding: utf-8 -*-
"""Гейт 1: в базе нет Конфигуратора (он блокирует пакетную загрузку, а со старой копией — откатывает её при F7)."""
import collections
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import connect  # noqa: E402

B = connect()
arr = B.ПолучитьСеансыИнформационнойБазы()
apps = collections.Counter()
designer = []
for i in range(arr.Количество()):
    s = arr.Получить(i)
    app = B.String(s.ИмяПриложения)
    apps[app] += 1
    if app.lower() in ("designer", "конфигуратор", "конфігуратор"):
        designer.append((B.String(s.Пользователь), B.String(s.ИмяКомпьютера), B.String(s.НачалоСеанса)))
print("сеансов:", arr.Количество(), dict(apps))
if designer:
    print("ОТКРЫТ КОНФИГУРАТОР — НЕ ГРУЗИТЬ:", designer)
    sys.exit(1)
print("Конфигуратора нет.")
