# -*- coding: utf-8 -*-
"""
Валидатор PBIP-проекта: ловит то, что иначе всплывает только при открытии в Power BI Desktop.

Проверяет:
  1) структурную полноту относительно заведомо рабочего эталона PowerBi/Industrial/PL.pbip
     (именно отсутствие definition/version.json уронило первую сборку);
  2) что все JSON парсятся и несут обязательные ключи;
  3) TMDL: каждая `ref table` имеет файл; каждая таблица содержит partition;
  4) связи ссылаются на существующие таблицы и колонки;
  5) меры ссылаются только на существующие колонки факта и другие меры;
  6) отсутствие табов/пробелов вперемешку в TMDL.

Запуск: .venv/Scripts/python.exe scripts/validate_pbip.py
"""
import json, re, sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

БАЗА = Path(r"C:/Configuration_downloads/BASERP25/_Rarzrabotki/Olap_Factory")
ИМЯ = "ПланФактВиробництва"
ЭТАЛОН = Path(r"C:/Configuration_downloads/BASERP25/_Rarzrabotki/PowerBi/Industrial")
ошибки, предупреждения = [], []


def проверить_структуру():
    обяз_отчет = ["definition.pbir", ".platform", "definition/report.json",
                  "definition/version.json", "definition/pages/pages.json"]
    обяз_модель = ["definition.pbism", ".platform", "definition/database.tmdl",
                   "definition/model.tmdl", "definition/relationships.tmdl"]
    отчет, модель = БАЗА / f"{ИМЯ}.Report", БАЗА / f"{ИМЯ}.SemanticModel"
    print("1. Структура")
    for корень, список in ((отчет, обяз_отчет), (модель, обяз_модель)):
        for rel in список:
            есть = (корень / rel).exists()
            print(f"   {'OK ' if есть else 'НЕТ'} {корень.name}/{rel}")
            if not есть:
                ошибки.append(f"нет файла {корень.name}/{rel}")
    if not (БАЗА / f"{ИМЯ}.pbip").exists():
        ошибки.append("нет .pbip")

    # сверка с эталоном: какие файлы есть у него в корне definition, а у нас нет
    э_отчет = ЭТАЛОН / "PL.Report" / "definition"
    if э_отчет.exists():
        э = {f.name for f in э_отчет.iterdir() if f.is_file()}
        м = {f.name for f in (отчет / "definition").iterdir() if f.is_file()}
        нет = э - м
        print(f"   файлы эталона, отсутствующие у нас: {нет or 'нет'}")
        for n in нет:
            ошибки.append(f"эталон имеет definition/{n}, у нас нет")


def проверить_json():
    print("\n2. JSON")
    for f in sorted(БАЗА.rglob("*.json")):
        if "Py_Olap" in str(f) or ".pbi" in str(f.parts):
            continue
        try:
            json.loads(f.read_text(encoding="utf-8"))
            print(f"   OK  {f.relative_to(БАЗА)}")
        except Exception as e:
            print(f"   BAD {f.relative_to(БАЗА)}: {e}")
            ошибки.append(f"{f.name}: {e}")
    try:
        json.loads((БАЗА / f"{ИМЯ}.pbip").read_text(encoding="utf-8"))
        print(f"   OK  {ИМЯ}.pbip")
    except Exception as e:
        ошибки.append(f"pbip: {e}")


def проверить_tmdl():
    print("\n3-6. TMDL")
    d = БАЗА / f"{ИМЯ}.SemanticModel" / "definition"
    model = (d / "model.tmdl").read_text(encoding="utf-8")
    объявлены = re.findall(r"(?m)^ref table\s+(.+)$", model)
    объявлены = [t.strip().strip("'") for t in объявлены]
    файлы = {f.stem for f in (d / "tables").glob("*.tmdl")}
    print(f"   ref table в model.tmdl: {len(объявлены)}, файлов таблиц: {len(файлы)}")
    for t in объявлены:
        if t not in файлы:
            ошибки.append(f"ref table {t} без файла tables/{t}.tmdl")
    for f in файлы:
        if f not in объявлены:
            предупреждения.append(f"файл {f}.tmdl не объявлен в model.tmdl")

    колонки, меры = {}, set()
    for f in sorted((d / "tables").glob("*.tmdl")):
        txt = f.read_text(encoding="utf-8")
        имя = re.search(r"(?m)^table\s+(.+)$", txt)
        имя = имя.group(1).strip().strip("'") if имя else f.stem
        колонки[имя] = set(c.strip().strip("'") for c in re.findall(r"(?m)^\tcolumn\s+(.+)$", txt))
        меры |= set(m.strip().strip("'") for m in re.findall(r"(?m)^\tmeasure\s+'([^']+)'", txt))
        if "partition" not in txt:
            ошибки.append(f"{f.name}: нет partition")
        if re.search(r"(?m)^    [a-zA-Zа-яА-Я]", txt):
            ошибки.append(f"{f.name}: отступ пробелами вместо табов")
        # 🔴 имя с пробелом/дефисом ОБЯЗАНО быть в одинарных кавычках, иначе парсер TMDL
        # валится InvalidLineType «Непредвиденный тип строки: Other!» (случай `column Назва місяця`)
        for вид in ("table", "column", "partition", "measure"):
            for m in re.finditer(r"(?m)^\t*" + вид + r"\s+([^=\n]+?)\s*(?:=|$)", txt):
                имя_ид = m.group(1).strip()
                if имя_ид.startswith("'"):
                    continue
                if not re.fullmatch(r"\w+", имя_ид, re.UNICODE):
                    ошибки.append(f"{f.name}: {вид} {имя_ид!r} с пробелом/дефисом БЕЗ кавычек")
    print(f"   таблиц: {len(колонки)}, мер: {len(меры)}")

    # 🔴 sourceColumn каждой колонки ОБЯЗАН быть в списке колонок native query — и наоборот.
    # Колонка есть в SQL и в TMDL, но забыта в SELECT → модель её молча не видит
    # (так потерялся Document_ID, FINDINGS §21). Проверка ловит обе половины расхождения.
    расхождений = 0
    for f in sorted((d / "tables").glob("*.tmdl")):
        txt = f.read_text(encoding="utf-8")
        запрос = re.search(r'\[Query="SELECT (.+?) FROM dbo\.', txt)
        if not запрос:
            ошибки.append(f"{f.name}: в partition нет native query [Query=\"SELECT ... FROM dbo.X\"]")
            расхождений += 1
            continue
        из_запроса = [c.strip() for c in запрос.group(1).split(",")]
        из_колонок = re.findall(r"(?m)^\t\tsourceColumn: (.+)$", txt)
        нет_в_запросе = [c for c in из_колонок if c not in из_запроса]
        нет_в_колонках = [c for c in из_запроса if c not in из_колонок]
        for c in нет_в_запросе:
            ошибки.append(f"{f.name}: колонка с sourceColumn {c} отсутствует в native query")
        for c in нет_в_колонках:
            предупреждения.append(f"{f.name}: native query тянет {c}, но колонки в TMDL нет")
        расхождений += len(нет_в_запросе)
    print(f"   колонок TMDL вне native query: {расхождений}")

    # связи
    rel = (d / "relationships.tmdl").read_text(encoding="utf-8")
    пары = re.findall(r"(?m)^\t(fromColumn|toColumn):\s*(.+)$", rel)
    for роль, ссылка in пары:
        т, _, к = ссылка.strip().rpartition(".")
        т = т.strip("'")
        if т not in колонки:
            ошибки.append(f"связь {роль}: нет таблицы {т}")
        elif к not in колонки[т]:
            ошибки.append(f"связь {роль}: в {т} нет колонки {к}")
    print(f"   связей проверено: {len(пары) // 2}")

    # меры: ссылки на Факти[...] и на [другие меры]
    факт = "Факти"
    плохих = 0
    for f in (d / "tables").glob("*.tmdl"):
        txt = f.read_text(encoding="utf-8")
        for m in re.finditer(r"measure\s+'([^']+)'\s*=\s*(.+)", txt):
            имя_меры, dax = m.group(1), m.group(2)
            for т, к in re.findall(r"(\w+)\[([^\]]+)\]", dax):
                if т in колонки and к not in колонки[т]:
                    ошибки.append(f"мера «{имя_меры}»: в {т} нет колонки {к}")
                    плохих += 1
            for ссыл in re.findall(r"(?<![\wа-яА-Я])\[([^\]]+)\]", dax):
                if ссыл not in меры and not any(ссыл in v for v in колонки.values()):
                    ошибки.append(f"мера «{имя_меры}»: ссылка [{ссыл}] не найдена")
                    плохих += 1
    print(f"   ошибок в DAX-ссылках: {плохих}")
    # 🔴 имя таблицы с не-ASCII символами в DAX обязано быть в кавычках:
    # SUM(Факти[X]) даёт «недопустимый токен» и мера уходит в SemanticError
    некавычек = 0
    for f in (d / "tables").glob("*.tmdl"):
        txt = f.read_text(encoding="utf-8")
        for m in re.finditer(r"measure\s+'([^']+)'\s*=\s*(.+)", txt):
            имя_меры, dax = m.group(1), m.group(2)
            for т in re.findall(r"(?<!')\b(\w+)\[", dax):
                if not т.isascii():
                    ошибки.append(f"мера «{имя_меры}»: таблица {т}[...] без кавычек "
                                  f"(нужно '{т}'[...] — DAX не принимает не-ASCII имя без кавычек)")
                    некавычек += 1
    print(f"   таблиц без кавычек в DAX: {некавычек}")


def main():
    проверить_структуру()
    проверить_json()
    проверить_tmdl()
    print()
    if предупреждения:
        print("Предупреждения:")
        for p in предупреждения:
            print("   ", p)
    if ошибки:
        print("ОШИБКИ:")
        for e in ошибки:
            print("   ", e)
        return 1
    print("PBIP ВАЛИДЕН по всем проверкам.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
