# -*- coding: utf-8 -*-
"""Rule #4 при отставшем worktree (Д1): копия WORKTREE → ОСНОВНОЙ каталог дампа BuhBud.

- файл, который уже равен основному (md5 без CR), не копируется: «без изменений» (корни ролей в списке e1a не правятся);
- файл из манифеста синхронизации: копируется, только если основной не менялся (md5 = манифест); иначе отказ —
  выполнить sync_main_to_wt.py --file <rel>, повторить свой патч-скрипт (идемпотентен) и to_main.py;
- новый файл П1 (нет в манифесте): копируется, если в основном его нет или он уже равен копии worktree;
- после копии md5 сверяется, манифест обновляется.
Запуск:
    python to_main.py <rel> [<rel> …]          # копировать
    python to_main.py --list <файл списка>     # копировать все файлы списка, которые есть в worktree
    python to_main.py --verify <файл списка>   # 0 — каждый файл списка в worktree совпадает с основным
Корневой файл формы в списке тянет её Ext/Form.xml и Ext/Form/Module.bsl (загрузка берёт форму целиком) — только те,
что П1 правит: новые формы П1 и синхронизированные (есть в манифесте). Устаревший модуль формы из worktree не копируется.
"""
import argparse
import json
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, BB_WT, HERE, md5_nocr  # noqa: E402

MANIFEST = os.path.join(HERE, "sync_manifest.json")


def read_list(path):
    return [ln.strip() for ln in open(path, encoding="utf-8-sig") if ln.strip() and not ln.startswith("#")]


def is_new_p1(rel):
    return any(m in rel for m in ("А_СпецификацииДомов", "А_ВидыСпецификацийДомов", "А_СтатусыСпецификацийДомов",
                                  "А_ВидыСтрокСпецификацииДомов", "А_ИсточникиКарточкиСпецификации",
                                  "А_ГруппыЗакупочнойСпецификации", "А_ИсточникиЦеныЗакупки"))


def expand(rels):
    manifest = json.load(open(MANIFEST, encoding="utf-8")) if os.path.exists(MANIFEST) else {}
    out = []
    for rel in rels:
        out.append(rel)
        m = re.match(r"(.*/Forms/[^/]+)\.xml$", rel)
        if m:
            for part in (m.group(1) + "/Ext/Form.xml", m.group(1) + "/Ext/Form/Module.bsl"):
                if os.path.exists(os.path.join(BB_WT, part)) and (part in manifest or is_new_p1(part)):
                    out.append(part)
    return out


def copy(rels):
    manifest = json.load(open(MANIFEST, encoding="utf-8"))
    fails = []
    for rel in rels:
        wt, mn = os.path.join(BB_WT, rel), os.path.join(BB_MAIN, rel)
        if not os.path.exists(wt):
            fails.append(f"{rel}: нет в worktree")
            continue
        if os.path.exists(mn) and md5_nocr(mn) == md5_nocr(wt):
            print(f"  без изменений {rel}")
            continue
        if rel in manifest:
            if os.path.exists(mn) and md5_nocr(mn) != manifest[rel] and md5_nocr(mn) != md5_nocr(wt):
                fails.append(f"{rel}: основной каталог изменён после синхронизации — sync_main_to_wt.py --file {rel}, "
                             f"повторить патч, затем to_main.py")
                continue
        elif not is_new_p1(rel):
            fails.append(f"{rel}: не в манифесте и не новый объект П1 — сначала синхронизировать")
            continue
        elif os.path.exists(mn) and md5_nocr(mn) != md5_nocr(wt):
            print(f"  ПЕРЕЗАПИСЬ нового файла П1 в основном каталоге: {rel}")
        os.makedirs(os.path.dirname(mn), exist_ok=True)
        shutil.copyfile(wt, mn)
        assert md5_nocr(mn) == md5_nocr(wt), rel
        manifest[rel] = md5_nocr(mn)
        print(f"  скопирован {rel}")
    with open(MANIFEST, "w", encoding="utf-8", newline="\n") as f:   # формат sync_main_to_wt.py: LF и \n в конце
        f.write(json.dumps(manifest, ensure_ascii=False, indent=1, sort_keys=True) + "\n")
    for f in fails:
        print("ОТКАЗ:", f)
    return 1 if fails else 0


def verify(rels):
    bad = []
    for rel in rels:
        wt, mn = os.path.join(BB_WT, rel), os.path.join(BB_MAIN, rel)
        if not os.path.exists(mn):
            bad.append(f"{rel}: нет в основном каталоге")
        elif os.path.exists(wt) and md5_nocr(wt) != md5_nocr(mn):
            bad.append(f"{rel}: worktree ≠ основной (to_main.py не выполнен?)")
    for b in bad:
        print("НЕ СОВПАДАЕТ:", b)
    print(f"файлов {len(rels)}, расхождений {len(bad)}")
    return 1 if bad else 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("rels", nargs="*")
    ap.add_argument("--list")
    ap.add_argument("--verify")
    a = ap.parse_args()
    if a.verify:
        sys.exit(verify(expand(read_list(a.verify))))
    rels = a.rels + ([r for r in expand(read_list(a.list)) if os.path.exists(os.path.join(BB_WT, r))] if a.list else [])
    sys.exit(copy(rels))
