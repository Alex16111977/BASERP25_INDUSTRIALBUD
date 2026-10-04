# -*- coding: utf-8 -*-
"""Сборка модулей П1 из исходников spec_domov/bsl/*.bsl в дамп WORKTREE (UTF-8 BOM + CRLF).
Строка «<отступ>//@ЗАПРОС <имя>» заменяется многострочным литералом BSL из zaprosy/<имя>.sql (кавычки удваиваются,
строки после первой начинаются с «|»): в BSL попадает тот же текст, что исполнил Python COM тест (Rule #-1).
Партии накопительные: 1 → 2a → 2b → 2c. Запуск: python build_bsl.py --partiya 1"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_WT, HERE, write_text, zapros  # noqa: E402

ORDER = ["1", "2a", "2b", "2c"]
SERVER_PARTS = [("1", "srv_10_obschee.bsl"), ("1", "srv_15_edinicy.bsl"), ("1", "srv_20_derevo.bsl"),
                ("1", "srv_30_cikly.bsl"), ("1", "srv_32_proverki.bsl"), ("1", "srv_35_etalon.bsl"),
                ("2a", "srv_40_edinicy2.bsl"), ("2a", "srv_50_fakt.bsl"), ("2a", "srv_55_zapolnenie.bsl"),
                ("2b", "srv_60_uzly.bsl"), ("2b", "srv_70_nulevye.bsl"), ("2b", "srv_80_ceny.bsl"),
                ("2b", "srv_90_sravnenie.bsl"), ("2c", "srv_A0_perepriv.bsl"), ("2c", "srv_B0_kopiya.bsl")]
HEADER = ("////////////////////////////////////////////////////////////////////////////////\n"
          "// А_СпецификацииДомовСервер — многоуровневая спецификация домов и узлов (П1 «Специфікація і еталон»).\n"
          "// Спецификация: docs/superpowers/specs/2026-10-03-mnogourovnevaya-specifikaciya-domov-design.md (§4–§6).\n"
          "// Сервер + внешнее соединение, не привилегированный. Собран build_bsl.py из _Rarzrabotki/Python/test/spec_domov/bsl;\n"
          "// тексты запросов — spec_domov/zaprosy/*.sql (проверены Python COM). Править исходники, а не этот файл.\n"
          "////////////////////////////////////////////////////////////////////////////////\n\n")
MODULE = "CommonModules/А_СпецификацииДомовСервер/Ext/Module.bsl"
OBJECT_MODULE = "Catalogs/А_СпецификацииДомов/Ext/ObjectModule.bsl"
RE_MARK = re.compile(r"^(?P<ind>[ \t]*)//@ЗАПРОС (?P<name>\w+)[ \t]*$", re.M)


def literal(name, ind):
    lines = zapros(name).replace('"', '""').split("\n")
    out = [f'{ind}"{lines[0]}'] + [f"{ind}|{ln}" for ln in lines[1:]]
    out[-1] += '"'
    return "\n".join(out)


def expand(text):
    text = RE_MARK.sub(lambda m: literal(m.group("name"), m.group("ind")), text)
    assert "//@ЗАПРОС" not in text
    return text


def read_src(name):
    return open(os.path.join(HERE, "bsl", name), encoding="utf-8-sig").read().replace("\r\n", "\n")


def build(partiya):
    upto = ORDER.index(partiya)
    parts = [read_src(f) for p, f in SERVER_PARTS if ORDER.index(p) <= upto]
    text = HEADER + "\n".join(expand(p).rstrip("\n") + "\n" for p in parts)
    write_text(os.path.join(BB_WT, MODULE), text)
    print(f"  {MODULE}: частей {len(parts)}, строк {text.count(chr(10))}")
    # Модуль объекта грузится только партией 1B: в партиях 2A–2C его не трогать, иначе worktree-копия разойдётся
    # с каноном платформы в основном каталоге (dump_check --canon, Task 28) и to_main.py --verify e1b (Task 45) упадёт.
    if partiya == "1" and os.path.exists(os.path.join(HERE, "bsl", "obj_module.bsl")):
        write_text(os.path.join(BB_WT, OBJECT_MODULE), expand(read_src("obj_module.bsl")))
        print(f"  {OBJECT_MODULE}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--partiya", choices=ORDER, required=True)
    build(ap.parse_args().partiya)
