# -*- coding: utf-8 -*-
"""Патчер: захват источника ПрочиеАктивыПассивы в Свод_ПрочиеАктивыПассивы_Прямой
и Свод_ОплатаТруда. По индексам строк с проверкой содержимого (assert).
Применяется к worktree-копии ObjectModule.bsl."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

P = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\angry-poincare-d9bb04\Documents\А_ФинРез_Баланс\Ext\ObjectModule.bsl"
lines = open(P, encoding="utf-8", newline="").readlines()
assert len(lines) == 1654, f"неожиданно строк: {len(lines)}"

PREF = "\t|\t"
NL = "\r\n"

FILTER_OLD = PREF + "И Т.Источник = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.ПустаяСсылка)" + NL
FILTER_NEW = (PREF + "И Т.Источник В (ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.ПустаяСсылка), "
              "ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.ПрочиеАктивыПассивы))" + NL)
SOURCE_OLD = PREF + "Т.Источник КАК Source," + NL
SOURCE_NEW = PREF + "ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.ПустаяСсылка) КАК Source," + NL
STAT_COMMA = PREF + "Т.Статья," + NL
STAT_NOCOMMA = PREF + "Т.Статья" + NL
IST_GROUP = PREF + "Т.Источник" + NL

# (1-based) -> 0-based
def idx(n): return n - 1

# проверки до правки
for n in (1188, 1231, 1407):
    assert lines[idx(n)] == FILTER_OLD, f"строка {n} не фильтр: {lines[idx(n)]!r}"
for n in (1218, 1394):
    assert lines[idx(n)] == SOURCE_OLD, f"строка {n} не Source: {lines[idx(n)]!r}"
for n in (1239, 1413):
    assert lines[idx(n)] == STAT_COMMA, f"строка {n} не 'Т.Статья,': {lines[idx(n)]!r}"
for n in (1240, 1414):
    assert lines[idx(n)] == IST_GROUP, f"строка {n} не 'Т.Источник' (группировка): {lines[idx(n)]!r}"

# применяем (порядок не важен — модификации по индексу, удаление через флаг)
for n in (1188, 1231, 1407):
    lines[idx(n)] = FILTER_NEW
for n in (1218, 1394):
    lines[idx(n)] = SOURCE_NEW
for n in (1239, 1413):
    lines[idx(n)] = STAT_NOCOMMA           # убрать запятую у Т.Статья
DELETE = {idx(1240), idx(1414)}            # удалить строки группировки Т.Источник
lines = [ln for i, ln in enumerate(lines) if i not in DELETE]

open(P, "w", encoding="utf-8", newline="").writelines(lines)
print(f"OK: фильтр×3, Source×2, группировка×2 (удалено 2 строки). Новых строк: {len(lines)}")
