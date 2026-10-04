# -*- coding: utf-8 -*-
"""check_static.py ловит типовые грабли BSL до загрузки и пропускает чистый модуль."""
import os
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import HERE, write_text  # noqa: E402

GOOD = """#Область Тест

Функция Проба(Знач Текст) Экспорт
	Запрос = Новый Запрос(
	"ВЫБРАТЬ
	|	Т.Ссылка КАК Ссылка
	|ИЗ
	|	Справочник.Номенклатура КАК Т");
	Если Текст = "" Тогда
		Возврат НСтр("ru = 'Пусто'; uk = 'Порожньо'");
	КонецЕсли;
	Возврат Запрос;
КонецФункции

#КонецОбласти
"""
BAD = (GOOD.replace("КонецФункции", "КонецФункция").replace("КАК Т\"", "КАК ПО\"").replace("uk = 'Порожньо'", "")
       .replace("\tЕсли Текст", "\tДокументы = Новый Массив;\n\tЕсли Текст"))


def run(path):
    return subprocess.run([sys.executable, "-X", "utf8", os.path.join(HERE, "check_static.py"), path],
                          capture_output=True, text=True, encoding="utf-8")


if __name__ == "__main__":
    d = tempfile.mkdtemp()
    good, bad = os.path.join(d, "good.bsl"), os.path.join(d, "bad.bsl")
    write_text(good, GOOD)
    write_text(bad, BAD)
    r = run(good)
    assert r.returncode == 0, r.stdout
    r = run(bad)
    assert r.returncode == 1, r.stdout
    for needle in ("КонецФункция", "КАК ПО", "НСтр без uk", "баланс Функция", "«Документы»"):
        assert needle in r.stdout, (needle, r.stdout)
    print("OK test_check_static")
