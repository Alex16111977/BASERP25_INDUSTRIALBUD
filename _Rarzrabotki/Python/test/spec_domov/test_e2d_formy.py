# -*- coding: utf-8 -*-
"""Формы этапа 2. --files (до загрузки): у каждой команды формы — <Command> с <Action> и обработчик в модуле; у каждого
ОписаниеОповещения("X", …) — экспортная процедура X; имена процедур модуля не повторяются; id элементов уникальны;
формы подбора и перепривязки — параметры, таблицы, события. Без ключа (после загрузки 2D) — пять форм справочника в базе."""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_WT, connect, read_text  # noqa: E402

FORMS = "Catalogs/А_СпецификацииДомов/Forms"
COMMANDS = {
    "ФормаЭлемента": ["СделатьЭталономСС", "ПоказатьВходимость", "ОбновитьДерево", "РазвернутьДоДеталей", "ЗаполнитьИзСС",
                      "ОтметитьЗеленыеПроверенными", "ПодобратьКарточку", "ВзятьИзОригиналаExcel", "ОбъединитьВУзел",
                      "РасформироватьУзел", "СоставИзДокументаКомплектации", "РассчитатьЦеныНормы", "СравнитьСоСС",
                      "НулевыеПозицииСФактом", "СкопироватьИзДругойСпецификации"],
    "ФормаСписка": ["ПроверитьЦиклы", "ПерепривязатьПозицииСС"],
    "ФормаПодбораКарточки": ["ВыбратьКарточку", "ОбновитьКандидатов"],
    "ФормаПерепривязки": ["НайтиПредложения", "ПрименитьОтмеченные"],
}
EVENTS = {
    "ФормаЭлемента": ["ПриСозданииНаСервере", "ВидСпецификацииПриИзменении", "ТолькоПроблемныеПриИзменении",
                      "СоставНоменклатураПриИзменении"],
    "ФормаПодбораКарточки": ["ПриСозданииНаСервере", "ТолькоОНПриИзменении", "СтрокаПоискаПриИзменении", "КандидатыВыбор"],
    "ФормаПерепривязки": ["ПриСозданииНаСервере", "ТипБудинкуССПриИзменении"],
}
TABLES = {"ФормаПодбораКарточки": "Кандидаты", "ФормаПерепривязки": "Предложения"}
PARAMS = {"ФормаПодбораКарточки": ["ТипБудинкуСС", "ОбщееНазвание", "ПозицияСС"], "ФормаПерепривязки": ["ТипБудинкуСС"]}


def files(form):
    x = read_text(os.path.join(BB_WT, FORMS, form, "Ext", "Form.xml"))[0]
    m = read_text(os.path.join(BB_WT, FORMS, form, "Ext", "Form", "Module.bsl"))[0]
    return x, m


def proc_names(module):
    return re.findall(r"(?m)^(?:Процедура|Функция) (\w+)\(", module)


def check_files():
    fails = []
    for form, cmds in COMMANDS.items():
        try:
            x, m = files(form)
        except FileNotFoundError as e:
            fails.append(f"{form}: нет файла {e.filename}")
            continue
        names = proc_names(m)
        dup = sorted({n for n in names if names.count(n) > 1})
        if dup:
            fails.append(f"{form}: повтор процедур {dup}")
        for c in cmds:
            if f'<Command name="{c}"' not in x or f"<Action>{c}</Action>" not in x:
                fails.append(f"{form}: нет команды {c}")
            if not re.search(r"(?m)^Процедура " + c + r"\(Команда\)", m):
                fails.append(f"{form}: нет обработчика команды {c}")
        for ev in EVENTS.get(form, []):
            if not re.search(r"(?m)^Процедура " + ev + r"\(", m):
                fails.append(f"{form}: нет обработчика события {ev}")
        for cb in sorted(set(re.findall(r'ОписаниеОповещения\("(\w+)"', m))):
            if not re.search(r"(?m)^Процедура " + cb + r"\([^)]*\) Экспорт", m):
                fails.append(f"{form}: оповещение {cb} без экспортной процедуры")
        ids = re.findall(r' id="(\d+)"', x.split("<Attributes>")[0])
        if len(ids) != len(set(ids)):
            fails.append(f"{form}: повтор id элементов")
        if form in TABLES and f'<Table name="{TABLES[form]}"' not in x:
            fails.append(f"{form}: нет таблицы {TABLES[form]}")
        for p in PARAMS.get(form, []):
            if f'<Parameter name="{p}"' not in x:
                fails.append(f"{form}: нет параметра {p}")
    return fails


def check_com():
    B = connect()
    cat = B.Metadata.Catalogs.Find("А_СпецификацииДомов")
    have = [cat.Forms.Get(i).Name for i in range(cat.Forms.Count())]
    fails = [f"нет формы {f} в базе" for f in ("ФормаЭлемента", "ФормаСписка", "ФормаГруппы", "ФормаПодбораКарточки",
                                                 "ФормаПерепривязки") if f not in have]
    for prop, form in (("DefaultObjectForm", "ФормаЭлемента"), ("DefaultListForm", "ФормаСписка"),
                       ("DefaultFolderForm", "ФормаГруппы")):
        if getattr(cat, prop).Name != form:
            fails.append(f"{prop} = {getattr(cat, prop).Name}")
    return fails


if __name__ == "__main__":
    fails = check_files() if "--files" in sys.argv else check_com()
    for f in fails:
        print("FAIL:", f)
    print(f"ПРОВАЛОВ: {len(fails)}")
    sys.exit(1 if fails else 0)
