# -*- coding: utf-8 -*-
"""to_main.py: копирует worktree → основной только при неизменном основном (хеш манифеста); --verify сверяет список."""
import os
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import HERE  # noqa: E402

PY = [sys.executable, "-X", "utf8", os.path.join(HERE, "to_main.py")]


def test_verify_synced_list():
    lst = os.path.join(tempfile.mkdtemp(), "l.txt")
    open(lst, "w", encoding="utf-8").write("Configuration.xml\nDocuments/РасчетКомплектаций.xml\nRoles/ЧтениеДанныхБухгалтерии.xml\n")
    r = subprocess.run(PY + ["--verify", lst], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stdout + r.stderr


def test_refuse_unknown_existing():
    """Номенклатура.xml есть в обоих каталогах и РАЗЛИЧАЕТСЯ (worktree отстал), в манифесте её нет → отказ."""
    r = subprocess.run(PY + ["Catalogs/Номенклатура.xml"], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 1 and "не в манифесте" in r.stdout, r.stdout


def test_skip_unchanged():
    """Файл списка, равный основному (корень роли), не копируется и не даёт отказа — иначе список e1a падал бы на Roles/X.xml."""
    r = subprocess.run(PY + ["Roles/ЧтениеДанныхБухгалтерии.xml"], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0 and "без изменений" in r.stdout, r.stdout


def test_expand_forms():
    """Корень формы СС тянет синхронизированный Ext/Form.xml, но не модуль формы, которого нет в манифесте."""
    import to_main
    root = "Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента.xml"
    out = to_main.expand([root])
    assert out == [root, "Catalogs/СтруктураСебестоимости/Forms/ФормаЭлемента/Ext/Form.xml"], out


if __name__ == "__main__":
    test_verify_synced_list()
    test_refuse_unknown_existing()
    test_skip_unchanged()
    test_expand_forms()
    print("OK test_to_main")
