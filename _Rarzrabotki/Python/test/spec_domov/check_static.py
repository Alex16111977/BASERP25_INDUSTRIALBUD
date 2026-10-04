# -*- coding: utf-8 -*-
"""Статический гейт П1 (E2 §4.1 П5).
XML: BOM, CRLF, well-formed, без ' />' и d5p1, у новых объектов П1 ru = uk; формы — уникальные id, <Action> без xsi:type.
BSL: BOM, CRLF, баланс Процедура/Функция/Если/Цикл/Попытка (без строк и комментариев), опечатки КонецФункция/КонецПроцедура,
псевдонимы-ключевые слова в запросах (КАК В/И/ПО…), SQLSERVER, Асинх/Ждать (совместимость 8.3.12), присваивание
зарезервированным именам и именам общих модулей (память feedback_vypolnit_reserved), НСтр без uk, неразвёрнутый //@ЗАПРОС,
вызов метода прямо на Новый; в модулях форм — директива перед каждой процедурой.
--tests: в spec_domov/*.py нет CommitTransaction/ЗафиксироватьТранзакцию/.Удалить().
Запуск: python check_static.py <путь|rel дампа>… | --list <список загрузки> | --tests"""
import argparse
import glob
import json
import os
import re
import sys
import xml.dom.minidom

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, BB_WT, HERE  # noqa: E402

RESERVED = {"Ссылка", "Метаданные", "Выполнить", "Знач", "ЭтотОбъект", "Строка", "Число", "Дата", "Тип", "Булево",
            # свойства глобального контекста: неявная локальная с таким именем падает «Поле объекта недоступно для записи»
            "Справочники", "Документы", "Перечисления", "Константы", "Отчеты", "Обработки", "ПланыВидовХарактеристик",
            "ПланыСчетов", "ПланыОбмена", "РегистрыСведений", "РегистрыНакопления", "РегистрыБухгалтерии",
            "ЖурналыДокументов", "БизнесПроцессы", "Задачи", "ПараметрыСеанса", "Символы", "ВнешниеОбработки", "ВнешниеОтчеты"}
FORM_MEMBERS = {"Параметры", "Элементы", "Объект", "Команды", "УсловноеОформление", "Модифицированность"}
KW = ("В", "И", "ПО", "ИЛИ", "НЕ", "ИЗ", "КАК", "ДЛЯ", "ГДЕ", "ВЫБРАТЬ", "ИМЕЮЩИЕ", "ИТОГИ", "МЕЖДУ", "ЕСТЬ",
      "ВЫБОР", "КОГДА", "ТОГДА", "ИНАЧЕ", "КОНЕЦ", "ПОДОБНО", "СОЕДИНЕНИЕ")
L = r"[A-Za-zА-Яа-яЁёІіЇїЄєҐґ_0-9]"
NEW_MARK = ("А_СпецификацииДомов", "А_ВидыСпецификацийДомов", "А_СтатусыСпецификацийДомов", "А_ВидыСтрокСпецификацииДомов",
            "А_ИсточникиКарточкиСпецификации", "А_ГруппыЗакупочнойСпецификации", "А_ИсточникиЦеныЗакупки")


def module_names():
    return {os.path.splitext(f)[0] for f in os.listdir(os.path.join(BB_MAIN, "CommonModules")) if f.endswith(".xml")}


def strip_code(text):
    """Текст модуля без строковых литералов (→ "") и комментариев; переводы строк сохраняются."""
    out, i, n, in_str = [], 0, len(text), False
    while i < n:
        c = text[i]
        if in_str:
            if c == '"':
                if i + 1 < n and text[i + 1] == '"':
                    i += 2
                    continue
                in_str = False
                out.append('""')
            elif c == "\n":
                out.append("\n")
            i += 1
            continue
        if text.startswith("//", i):
            j = text.find("\n", i)
            i = n if j < 0 else j
            continue
        if c == '"':
            in_str = True
            i += 1
            continue
        out.append(c)
        i += 1
    return "".join(out)


def count(pattern, code):
    return len(re.findall(pattern, code, re.M))


def check_raw(raw, errs, label):
    if not raw.startswith(b"\xef\xbb\xbf"):
        errs.append(f"{label}: нет BOM")
    if raw.replace(b"\r\n", b"").count(b"\n"):
        errs.append(f"{label}: не CRLF")


def check_bsl(label, raw, is_form, mods):
    errs = []
    check_raw(raw, errs, label)
    text = raw.decode("utf-8-sig").replace("\r\n", "\n")
    code = strip_code(text)
    pairs = [(r"^\s*Процедура\s", r"^\s*КонецПроцедуры\b", "Процедура"), (r"^\s*Функция\s", r"^\s*КонецФункции\b", "Функция"),
             (r"(?<!Иначе)(?<!" + L + r")Если(?!" + L + r")", r"(?<!" + L + r")КонецЕсли(?!" + L + r")", "Если"),
             (r"(?<!" + L + r")Цикл(?!" + L + r")", r"(?<!" + L + r")КонецЦикла(?!" + L + r")", "Цикл"),
             (r"(?<!" + L + r")Попытка(?!" + L + r")", r"(?<!" + L + r")КонецПопытки(?!" + L + r")", "Попытка")]
    for op, cl, nm in pairs:
        a, b = count(op, code), count(cl, code)
        if a != b:
            errs.append(f"{label}: баланс {nm}: {a} ≠ {b}")
    for m in re.finditer(r"КонецФункция\b|КонецПроцедура\b", code):
        errs.append(f"{label}: опечатка {m.group(0)}")
    for m in re.finditer(r"(?<!" + L + r")(Асинх|Ждать)(?!" + L + r")", code):
        errs.append(f"{label}: {m.group(1)} недоступно в режиме совместимости 8.3.12")
    if "SQLSERVER" in text:
        errs.append(f"{label}: SQLSERVER в коде (только localhost)")
    if "//@ЗАПРОС" in text and not label.startswith("src:"):
        errs.append(f"{label}: неразвёрнутый //@ЗАПРОС")
    for m in re.finditer(r"КАК\s+(" + "|".join(KW) + r")(?!" + L + r")", text):
        errs.append(f"{label}: псевдоним-ключевое слово «КАК {m.group(1)}»")
    targets = [m.group(1) for m in re.finditer(r"^[ \t]*(" + L + r"+)[ \t]*=(?!=)", code, re.M)]
    # переменная цикла — тоже неявная локальная («Для Каждого Документы Из …» падает так же, как присваивание)
    targets += [m.group(1) for m in re.finditer(r"^[ \t]*Для[ \t]+(?:Каждого[ \t]+)?(" + L + r"+)[ \t]*(?:Из(?!" + L + r")|=)",
                                                code, re.M | re.I)]
    for name in targets:
        if name in RESERVED or name in mods or (is_form and name in FORM_MEMBERS):
            errs.append(f"{label}: присваивание зарезервированному имени «{name}»")
    for m in re.finditer(r'НСтр\(\s*"((?:[^"]|"")*)"', text):
        if not re.search(r"\buk\s*=", m.group(1)):
            errs.append(f"{label}: НСтр без uk: {m.group(1)[:50]}")
    if re.search(r"Новый\s+" + L + r"+\([^()\n]*\)\.", code):
        errs.append(f"{label}: вызов метода прямо на Новый")
    if is_form:
        lines = text.split("\n")
        for i, line in enumerate(lines):
            if re.match(r"^(Процедура|Функция)\s", line):
                j = i - 1
                while j >= 0 and not lines[j].strip():
                    j -= 1
                if j < 0 or not lines[j].lstrip().startswith("&"):
                    errs.append(f"{label}: нет директивы компиляции перед «{line.strip()[:60]}»")
    return errs


def check_xml(label, raw, new_obj, is_form):
    errs = []
    check_raw(raw, errs, label)
    text = raw.decode("utf-8-sig").replace("\r\n", "\n")
    try:
        xml.dom.minidom.parseString(text.encode("utf-8"))
    except Exception as e:
        errs.append(f"{label}: не XML: {e}")
    if " />" in text or "d5p1" in text:
        errs.append(f"{label}: шум генератора (' />' или d5p1)")
    if new_obj and text.count("<v8:lang>ru</v8:lang>") != text.count("<v8:lang>uk</v8:lang>"):
        errs.append(f"{label}: ru {text.count('<v8:lang>ru</v8:lang>')} ≠ uk {text.count('<v8:lang>uk</v8:lang>')}")
    if is_form:
        if "<Action xsi:type" in text:
            errs.append(f"{label}: <Action> с xsi:type")
        body = text.split("<Attributes>")[0]
        ids = re.findall(r'<(?!Attribute|Command|Column|Parameter)\w+ name="[^"]*" id="(-?\d+)"', body)
        dup = {x for x in ids if ids.count(x) > 1}
        if dup:
            errs.append(f"{label}: повтор id элементов {sorted(dup)}")
        cids = re.findall(r'<Command name="[^"]*" id="(\d+)"', text)
        if len(cids) != len(set(cids)):
            errs.append(f"{label}: повтор id команд")
    return errs


def resolve(p):
    if os.path.isabs(p):
        return p
    if p.replace("\\", "/").startswith("_Rarzrabotki/Python/") and os.path.exists(p):
        return os.path.abspath(p)   # исходник spec_domov по пути от WT (команды Step 4 задач 38, 39, 42)
    wt = os.path.join(BB_WT, p)
    return wt if os.path.exists(wt) else os.path.join(BB_MAIN, p)


def expand(rels):
    out = []
    for r in rels:
        out.append(r)
        m = re.match(r"(.*/Forms/[^/]+)\.xml$", r)
        if m:
            out += [m.group(1) + "/Ext/Form.xml", m.group(1) + "/Ext/Form/Module.bsl"]
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="*")
    ap.add_argument("--list")
    ap.add_argument("--tests", action="store_true")
    a = ap.parse_args()
    errs = []
    if a.tests:
        for f in glob.glob(os.path.join(HERE, "*.py")):
            if os.path.basename(f) == "check_static.py":
                continue
            t = open(f, encoding="utf-8").read()
            for bad in ("CommitTransaction", "ЗафиксироватьТранзакцию", ".Удалить()"):
                if bad in t:
                    errs.append(f"{os.path.basename(f)}: запрещено в тестах: {bad}")
    rels = list(a.paths)
    if a.list:
        rels += [ln.strip() for ln in open(a.list, encoding="utf-8-sig") if ln.strip() and not ln.startswith("#")]
    mods = module_names()
    mf = os.path.join(HERE, "sync_manifest.json")
    manifest = json.load(open(mf, encoding="utf-8")) if os.path.exists(mf) else {}
    for rel in expand(rels):
        path = resolve(rel)
        if (rel.endswith("Module.bsl") and "/Forms/" in rel and not any(m in rel for m in NEW_MARK)
                and rel not in manifest):
            continue   # модуль существующей формы, который П1 не правит (как to_main.expand): живой код, не проверяем
        if not os.path.exists(path):
            if rel.endswith("Module.bsl") and "/Forms/" in rel:
                continue
            errs.append(f"{rel}: файла нет")
            continue
        raw = open(path, "rb").read()
        new_obj = any(m in rel for m in NEW_MARK)
        if path.endswith(".bsl"):
            errs += check_bsl(rel, raw, "/Forms/" in rel.replace("\\", "/"), mods)
        elif path.endswith(".xml") and "/Roles/" not in rel.replace("\\", "/") and not rel.endswith("Configuration.xml"):
            errs += check_xml(rel, raw, new_obj, rel.endswith("Form.xml"))
        elif path.endswith(".xml"):
            try:
                xml.dom.minidom.parseString(raw.decode("utf-8-sig").encode("utf-8"))
            except Exception as e:
                errs.append(f"{rel}: не XML: {e}")
    for e in errs:
        print("ОШИБКА:", e)
    print(f"проверено файлов: {len(expand(rels))}, ошибок: {len(errs)}")
    return 1 if errs else 0


if __name__ == "__main__":
    sys.exit(main())
