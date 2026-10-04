# -*- coding: utf-8 -*-
"""Новые объекты П1 навыком meta-compile: генерация в scratch-каталог с заглушкой Configuration.xml версии 2.13
(meta-compile переписывает Configuration.xml целиком — E1 §1.3), постпатч, копия XML объекта в дамп WORKTREE.
Configuration.xml дампа не трогается: регистрация — register.py. Пустые модули (только BOM) не копируются.
ОПАСНО перегенерировать уже загруженный объект: новые UUID = другой объект для платформы. Существующий файл в worktree
не перезаписывается (--force — только до первой загрузки).
Запуск: python gen_meta.py meta/e1_enums.json [--force]"""
import argparse
import datetime
import json
import os
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import postpatch  # noqa: E402
from common import BB_WT, HERE, SKILLS, TMP, write_text  # noqa: E402

STUB = ('<?xml version="1.0" encoding="UTF-8"?>\n<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses" version="2.13">'
        "<Configuration><ChildObjects/></Configuration></MetaDataObject>\n")
FOLDER = {"Catalog": "Catalogs", "Enum": "Enums", "CommonModule": "CommonModules"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("json")
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args()
    path = a.json if os.path.isabs(a.json) else os.path.join(HERE, a.json)
    defs = json.load(open(path, encoding="utf-8"))
    defs = defs if isinstance(defs, list) else [defs]
    scratch = os.path.join(TMP, "gen_meta", datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
    os.makedirs(scratch, exist_ok=True)
    # По одному объекту на вызов и свежая заглушка перед каждым: meta-compile читает начало Configuration.xml
    # Substring(0, min(2000, длина в БАЙТАХ)) — после BOM или кириллицы (её дописывает сам meta-compile) короткий файл
    # даёт ArgumentOutOfRange. Заглушка без BOM, ASCII — длина в байтах = длине в символах.
    for i, d in enumerate(defs):
        with open(os.path.join(scratch, "Configuration.xml"), "w", encoding="utf-8", newline="") as f:
            f.write(STUB)
        one = os.path.join(scratch, f"_def_{i}.json")
        with open(one, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False)
        r = subprocess.run(["powershell.exe", "-NoProfile", "-File",
                            os.path.join(SKILLS, r"meta-compile\scripts\meta-compile.ps1"),
                            "-JsonPath", one, "-OutputDir", scratch],
                           capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout[-800:])
        if r.returncode != 0:
            print(r.stderr[-2000:])
            return 1
    for d in defs:
        rel = f"{FOLDER[d['type']]}/{d['name']}.xml"
        src, dst = os.path.join(scratch, rel), os.path.join(BB_WT, rel)
        assert os.path.exists(src), f"meta-compile не создал {rel}"
        if os.path.exists(dst) and not a.force:
            print("  ПРОПУСК (уже есть — UUID не перегенерируем):", rel)
            continue
        postpatch.fix(src, catalog=(d["type"] == "Catalog"))
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copyfile(src, dst)
        print("  создан", rel)
    return 0


if __name__ == "__main__":
    sys.exit(main())
