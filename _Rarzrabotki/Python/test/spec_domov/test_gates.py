# -*- coding: utf-8 -*-
"""Гейты должны ЛОВИТЬ аварию: Configuration.xml без объекта базы → код 1; лог с «Объект удален» → код 1."""
import os
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import BB_MAIN, HERE, read_text, write_text  # noqa: E402

PY = [sys.executable, "-X", "utf8"]


def run(*args):
    return subprocess.run(PY + list(args), capture_output=True, text=True, encoding="utf-8").returncode


def test_config_gate_catches_lost_object():
    d = tempfile.mkdtemp()
    cfg = os.path.join(d, "Configuration.xml")
    text, _, _ = read_text(os.path.join(BB_MAIN, "Configuration.xml"))
    line = "\t\t\t<Catalog>СтруктураСебестоимости</Catalog>\n"
    assert text.count(line) == 1
    write_text(cfg, text.replace(line, ""))
    assert run(os.path.join(HERE, "preflight_config.py"), cfg) == 1, "гейт не поймал потерю СС"
    assert run(os.path.join(HERE, "preflight_config.py")) == 0, "основной каталог должен проходить"


def test_log_gate():
    d = tempfile.mkdtemp()
    bad, good = os.path.join(d, "bad.log"), os.path.join(d, "good.log")
    open(bad, "w", encoding="utf-8-sig").write("Загрузка...\nОбъект удален: Справочник.Х\n")
    open(good, "w", encoding="utf-8-sig").write("Загрузка конфигурации успешно завершена\n")
    assert run(os.path.join(HERE, "check_log.py"), "load", bad) == 1
    assert run(os.path.join(HERE, "check_log.py"), "load", good) == 0
    open(good, "w", encoding="utf-8-sig").write("Синтаксических ошибок не обнаружено!\n")
    assert run(os.path.join(HERE, "check_log.py"), "checkmodules", good, "0") == 0
    assert run(os.path.join(HERE, "check_log.py"), "checkmodules", good, "101") == 1


def test_sessions_and_roles():
    assert run(os.path.join(HERE, "sessions_gate.py")) == 0
    assert run(os.path.join(HERE, "preflight_roles.py")) == 0


if __name__ == "__main__":
    test_config_gate_catches_lost_object()
    test_log_gate()
    test_sessions_and_roles()
    print("OK test_gates")
