# -*- coding: utf-8 -*-
"""Обратная выгрузка объектов из базы во Temp и сверка с ОСНОВНЫМ каталогом (E2 §4.6).

Частичная выгрузка не включает формы и макеты — их перечислять (`Справочник.X.Форма.Y`).
Сравнение XML — по канонической форме C14N (порядок xmlns и атрибутов, пробелы между тегами, `<a/>` = `<a></a>` не важны).
--canon: файлы, отличающиеся только форматированием, заменяются выгруженной версией в основном каталоге И в worktree
(канон — версия платформы, память bsl_write_lf_crlf_canon). Различие по содержимому:
- у СУЩЕСТВУЮЩИХ объектов (СС, РК, форма СС, план обмена) — код 1 (загрузка не дошла или кеш Конфигуратора);
- у НОВЫХ файлов П1 (сгенерированы meta-compile/gen_forms) платформа сериализует загруженное по-своему (умолчания свойств):
  печатается начало различия, при --canon выгруженная версия становится каноном («ПРИНЯТО КАНОНОМ»), код 0. Что загрузка
  дошла, доказывают COM-тесты этапа (test_e1_metadata, test_e1_priemka, test_e2d_formy), а не побайтная сверка.
Запуск: python dump_check.py --tag e1a --objects "Справочник.А_СпецификацииДомов,Перечисление.А_ВидыСпецификацийДомов" [--canon]
"""
import argparse
import datetime
import difflib
import json
import os
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, BB_WT, HERE, SKILLS, TMP, V8_BIN, md5_nocr, read_text  # noqa: E402

MANIFEST = os.path.join(HERE, "sync_manifest.json")
NEW_MARK = ("А_СпецификацииДомов", "А_ВидыСпецификацийДомов", "А_СтатусыСпецификацийДомов", "А_ВидыСтрокСпецификацииДомов",
            "А_ИсточникиКарточкиСпецификации", "А_ГруппыЗакупочнойСпецификации", "А_ИсточникиЦеныЗакупки")


def normalized(path):
    text, _, _ = read_text(path)
    if path.endswith(".xml"):
        try:
            return ET.canonicalize(text, strip_text=True)
        except Exception:                                   # noqa: BLE001 — не XML: сравнить без пробелов между тегами
            return re.sub(r">\s+<", "><", text).strip()
    return "\n".join(line.rstrip() for line in text.split("\n")).strip()


def first_diff(a, b):
    """Первые строки различия канонических форм (для XML — после разбиения по тегам)."""
    sa, sb = (x.replace("><", ">\n<").split("\n") for x in (a, b))
    return "\n".join(list(difflib.unified_diff(sa, sb, "основной", "база", n=1, lineterm=""))[:12])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tag", required=True)
    ap.add_argument("--objects", required=True)
    ap.add_argument("--canon", action="store_true")
    a = ap.parse_args()
    out = os.path.join(TMP, f"dump_{a.tag}_{datetime.datetime.now():%Y%m%d_%H%M%S}")
    os.makedirs(out, exist_ok=True)
    cmd = ["powershell.exe", "-NoProfile", "-File", os.path.join(SKILLS, r"db-dump-xml\scripts\db-dump-xml.ps1"),
           "-V8Path", V8_BIN, "-InfoBaseServer", "localhost", "-InfoBaseRef", "bas_industrialbud",
           "-UserName", "cfo", "-Password", "2442", "-ConfigDir", out, "-Mode", "Partial", "-Objects", a.objects]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    print(r.stdout[-2000:])
    if r.returncode != 0:
        print(r.stderr[-2000:])
        return 1
    manifest = json.load(open(MANIFEST, encoding="utf-8")) if os.path.exists(MANIFEST) else {}
    same = fmt_only = 0
    content, new_content = [], []

    def adopt(src, rel, main_path, label):
        for dst in (main_path, os.path.join(BB_WT, rel)):
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copyfile(src, dst)
        manifest[rel] = md5_nocr(main_path)
        print(f"  {label}:", rel)

    for root, _, files in os.walk(out):
        for fn in files:
            if fn in ("Configuration.xml", "ConfigDumpInfo.xml"):
                continue
            src = os.path.join(root, fn)
            rel = os.path.relpath(src, out).replace("\\", "/")
            main_path = os.path.join(BB_MAIN, rel)
            if os.path.exists(main_path) and md5_nocr(main_path) == md5_nocr(src):
                same += 1
            elif os.path.exists(main_path) and normalized(main_path) == normalized(src):
                fmt_only += 1
                if a.canon:
                    adopt(src, rel, main_path, "канон")
            elif any(m in rel for m in NEW_MARK):
                new_content.append(rel)
                if os.path.exists(main_path):
                    print(f"   НОВЫЙ ФАЙЛ П1, платформа сериализовала иначе: {rel}\n{first_diff(normalized(main_path), normalized(src))}")
                if a.canon:
                    adopt(src, rel, main_path, "ПРИНЯТО КАНОНОМ (новый файл П1)")
            else:
                content.append(rel)
    if a.canon:
        with open(MANIFEST, "w", encoding="utf-8", newline="\n") as f:   # формат sync_main_to_wt.py
            f.write(json.dumps(manifest, ensure_ascii=False, indent=1, sort_keys=True) + "\n")
    print(f"выгрузка {out}: совпало {same}, только форматирование {fmt_only}, новых файлов П1 иначе {len(new_content)}, "
          f"различие по содержимому {len(content)}")
    for c in content:
        print("   СОДЕРЖИМОЕ РАЗЛИЧАЕТСЯ:", c)
    return 1 if content or ((fmt_only or new_content) and not a.canon) else 0


if __name__ == "__main__":
    sys.exit(main())
