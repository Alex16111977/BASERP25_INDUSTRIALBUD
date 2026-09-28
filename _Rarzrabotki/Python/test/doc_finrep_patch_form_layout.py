# -*- coding: utf-8 -*-
"""Идемпотентный пост-патч компоновки формы (то, чего не умеет form-compile).

Канон «без прокрутки формы» — скил bas-form-layout + knowledge_Бухгалтерия_РасчетКомплектаций §10:
на каждой странице ровно ОДИН вертикально-растягиваемый элемент (главная таблица без height),
всё прочее жёстко фиксировано. Компилятор эмитит только VerticalStretch=TRUE и не знает про
ChildItemsWidth, поэтому фиксацию мини-таблиц и равные доли ширины ставим здесь.

Что делает (всё — только если ещё не стоит):
  1. Autofill: убирает <Autofill>false</Autofill> у AutoCommandBar — иначе нет стандартных
     «Записать и закрыть»/«Записать» (граблю см. 08_forma_dokumenta.md).
  2. <VerticalStretch>false</VerticalStretch> мини-таблицам (перед <DataPath> — порядок
     свойств Table: … HeightInTableRows → VerticalStretch → Footer → DataPath).
  3. <ChildItemsWidth>Equal</ChildItemsWidth> горизонтальным группам с парой таблиц
     (порядок UsualGroup: … Representation → ChildItemsWidth → United → ShowTitle).
  4. <TitleLocation>None</TitleLocation> надписи-состоянию: Emit-LabelField компилятора
     titleLocation игнорирует, из-за чего на форме висела подпись «СостояниеРасчета».

Запускать ПОСЛЕ form-compile и ПЕРЕД doc_finrep_patch_uk.py.
"""
import re
import sys
import xml.dom.minidom

sys.stdout.reconfigure(encoding="utf-8")

ФОРМА = (r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh\Documents"
         r"\А_ФинансовыйОтчетПроизводства\Forms\ФормаДокумента\Ext\Form.xml")

# Таблицы с фиксированной высотой (height в JSON) — не должны растягиваться
ФИКСИРОВАННЫЕ_ТАБЛИЦЫ = ("ДоговорыДополнительно", "СчетаРасчетов", "СтатьиИсключенияКазны")
# Горизонтальные группы, дети которых делят ширину поровну
ГРУППЫ_РАВНОЙ_ШИРИНЫ = ("ГруппаУточнения",)
# Надписи без подписи-заголовка
НАДПИСИ_БЕЗ_ЗАГОЛОВКА = ("ПолеСостояние",)


def блок_элемента(текст, тег, имя):
    """Возвращает (начало, конец) XML-блока <Тег name="имя" …> … </Тег>."""
    начало = re.search(rf'<{тег} name="{re.escape(имя)}"[^>]*>', текст)
    if начало is None:
        return None
    конец = текст.find(f"</{тег}>", начало.end())
    if конец == -1:
        return None
    return начало.start(), конец + len(f"</{тег}>")


def вставить_перед(текст, границы, якорь, вставка, признак):
    """Вставляет `вставка` перед первым `якорь` внутри блока; идемпотентно по `признак`."""
    старт, стоп = границы
    блок = текст[старт:стоп]
    if признак in блок:
        return текст, False
    позиция = блок.find(якорь)
    if позиция == -1:
        return текст, False
    отступ = ""
    строка_начало = блок.rfind("\n", 0, позиция)
    if строка_начало != -1:
        отступ = блок[строка_начало + 1:позиция]
    новый_блок = блок[:позиция] + вставка + "\n" + отступ + блок[позиция:]
    return текст[:старт] + новый_блок + текст[стоп:], True


текст = open(ФОРМА, encoding="utf-8").read()
изменений = 0

# 1. Autofill
if "<Autofill>false</Autofill>" in текст:
    количество = текст.count("<Autofill>false</Autofill>")
    текст = текст.replace("<Autofill>false</Autofill>", "")
    изменений += количество
    print(f"Autofill=false убрано: {количество}")
else:
    print("Autofill: уже включён")

# 2. VerticalStretch=false мини-таблицам
for имя in ФИКСИРОВАННЫЕ_ТАБЛИЦЫ:
    границы = блок_элемента(текст, "Table", имя)
    if границы is None:
        print(f"[ПРОПУСК] таблица {имя} не найдена")
        continue
    текст, сделано = вставить_перед(текст, границы, "<DataPath>",
                                    "<VerticalStretch>false</VerticalStretch>",
                                    "<VerticalStretch>")
    изменений += сделано
    print(f"{имя}: VerticalStretch=false {'добавлен' if сделано else 'уже стоял'}")

# 3. ChildItemsWidth=Equal горизонтальным группам с таблицами
for имя in ГРУППЫ_РАВНОЙ_ШИРИНЫ:
    границы = блок_элемента(текст, "UsualGroup", имя)
    if границы is None:
        print(f"[ПРОПУСК] группа {имя} не найдена")
        continue
    якорь = "<ShowTitle>" if "<ShowTitle>" in текст[границы[0]:границы[1]] else "<ChildItems>"
    текст, сделано = вставить_перед(текст, границы, якорь,
                                    "<ChildItemsWidth>Equal</ChildItemsWidth>",
                                    "<ChildItemsWidth>")
    изменений += сделано
    print(f"{имя}: ChildItemsWidth=Equal {'добавлен' if сделано else 'уже стоял'}")

# 4. TitleLocation=None у надписи состояния
for имя in НАДПИСИ_БЕЗ_ЗАГОЛОВКА:
    границы = блок_элемента(текст, "LabelField", имя)
    if границы is None:
        print(f"[ПРОПУСК] надпись {имя} не найдена")
        continue
    текст, сделано = вставить_перед(текст, границы, "<ContextMenu",
                                    "<TitleLocation>None</TitleLocation>",
                                    "<TitleLocation>")
    изменений += сделано
    print(f"{имя}: TitleLocation=None {'добавлен' if сделано else 'уже стоял'}")

if изменений:
    open(ФОРМА, "w", encoding="utf-8").write(текст)
xml.dom.minidom.parse(ФОРМА)
print(f"\nПравок: {изменений}. XML валиден (minidom).")
