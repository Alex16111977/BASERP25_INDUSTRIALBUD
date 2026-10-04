# -*- coding: utf-8 -*-
"""Разбор логов загрузки (Tee-Object PowerShell 5.1 пишет UTF-16) и /CheckModules.
Запуск: check_log.py load <лог>  |  check_log.py checkmodules <лог> <код выхода 1cv8>. Код 0 — всё чисто."""
import sys

BAD_LOAD = ["Объект удален", "Объект удалён", "Неизвестный объект", "не входит в состав", "Неверное свойство",
            "нельзя добавлять", "Error loading configuration", "Ошибка загрузки"]
OK_CHECK = ["Синтаксических ошибок не обнаружено", "Синтаксичних помилок не виявлено"]


def read_log(path):
    raw = open(path, "rb").read()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16")
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw.decode("cp1251")


def main(argv):
    kind, path = argv[1], argv[2]
    text = read_log(path)
    if kind == "load":
        hits = [line for line in text.splitlines() if any(b in line for b in BAD_LOAD)]
        for h in hits:
            print("АВАРИЯ ЗАГРУЗКИ:", h)
        return 1 if hits else 0
    if kind == "checkmodules":
        code = int(argv[3])
        ok = code == 0 and any(o in text for o in OK_CHECK)
        if not ok:
            print(text[-5000:])
            print(f"CheckModules: код {code} — модули с ошибками (см. {{Объект(строка,колонка)}} и <<?>> выше)")
        return 0 if ok else 1
    raise SystemExit(f"неизвестный вид лога: {kind}")


if __name__ == "__main__":
    sys.exit(main(sys.argv))
