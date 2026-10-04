# -*- coding: utf-8 -*-
"""Д1: существующие файлы дампа, которые правит П1, копируются ОСНОВНОЙ → WORKTREE; хеши (без CR) пишутся в манифест.

Основной каталог = живая база (E2 §2), worktree-дамп отстал (E1 §0, E3 §0). to_main.py потом копирует правки
обратно, только если основной файл не менялся с момента синхронизации (хеш манифеста).
Запуск:
    python sync_main_to_wt.py --check             # 0 — worktree = основной по всем файлам SYNC
    python sync_main_to_wt.py                     # копировать все файлы SYNC и обновить манифест
    python sync_main_to_wt.py --file <rel>        # пересинхронизировать один файл (после отказа to_main.py)
"""
import argparse
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, BB_WT, HERE, md5_nocr  # noqa: E402

MANIFEST = os.path.join(HERE, "sync_manifest.json")
SYNC = [
    "Configuration.xml",
    "Catalogs/СтруктураСебестоимости.xml",
    "Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента.xml",
    "Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента/Ext/Form.xml",
    "Documents/РасчетКомплектаций.xml",
    "Roles/ДобавлениеИзменениеДанныхБухгалтерии/Ext/Rights.xml",
    "Roles/ЧтениеДанныхБухгалтерии/Ext/Rights.xml",
    "Roles/ЗапускТонкогоКлиента/Ext/Rights.xml",
    "Subsystems/Склад/Subsystems/А_СписанняБудинків.xml",
    "Subsystems/Склад/Subsystems/А_СписанняБудинків/Ext/CommandInterface.xml",
]


def load_manifest():
    if os.path.exists(MANIFEST):
        with open(MANIFEST, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_manifest(m):
    with open(MANIFEST, "w", encoding="utf-8", newline="\n") as f:
        json.dump(m, f, ensure_ascii=False, indent=1, sort_keys=True)
        f.write("\n")


def copy_one(rel, manifest):
    src, dst = os.path.join(BB_MAIN, rel), os.path.join(BB_WT, rel)
    assert os.path.exists(src), f"нет в основном каталоге: {rel}"
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copyfile(src, dst)                          # байт в байт: BOM и CRLF сохраняются
    h = md5_nocr(src)
    assert md5_nocr(dst) == h, rel
    manifest[rel] = h
    print(f"  синхронизирован {rel} md5={h}")


def check():
    bad = []
    for rel in SYNC:
        dst = os.path.join(BB_WT, rel)
        if not os.path.exists(dst) or md5_nocr(dst) != md5_nocr(os.path.join(BB_MAIN, rel)):
            bad.append(rel)
    for rel in bad:
        print("  РАСХОДИТСЯ:", rel)
    print(f"файлов {len(SYNC)}, расхождений {len(bad)}")
    return 1 if bad else 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--file")
    a = ap.parse_args()
    if a.check:
        sys.exit(check())
    m = load_manifest()
    for rel in ([a.file] if a.file else SYNC):
        copy_one(rel, m)
    save_manifest(m)
    sys.exit(check() if not a.file else 0)
