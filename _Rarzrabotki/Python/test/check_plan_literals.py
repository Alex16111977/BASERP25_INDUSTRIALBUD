# -*- coding: utf-8 -*-
"""Грабля 2: сверяет строки-константы планов обмена в ObjectModule.bsl с
ИМЕНАМИ ИЗ МЕТАДАННЫХ (plan_names.txt, получен из живой базы).

Проверяет:
  * литерал в Функция ПланБухгалтерии() == имя из метаданных;
  * литерал в Функция ПланКазны()       == имя из метаданных (буква «й» = 0x439!);
  * имя плана Казны встречается в файле РОВНО ОДИН РАЗ (требование ТЗ).

При расхождении — чинит файл байтами из метаданных (никакого набора руками).
"""
import io
import re
import sys

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

NAMES_FILE = "plan_names.txt"
TARGETS = [
    r"C:\Configuration_downloads\BASERP25\.claude\worktrees\erp-exchange-button-fa9a52"
    r"\_Rarzrabotki\Обработки\А_ЗапуститьОбмен\Ext\ObjectModule.bsl",
]

with io.open(NAMES_FILE, encoding="utf-8") as f:
    lines = [ln.strip() for ln in f if ln.strip()]
PLAN_BUH, PLAN_KAZNA = lines[0], lines[1]


def cps(text):
    return " ".join(hex(ord(c)) for c in text)


print(f"Из метаданных: Бухгалтерия = {PLAN_BUH}")
print(f"Из метаданных: Казна       = {PLAN_KAZNA}")
print(f"               codepoints  = {cps(PLAN_KAZNA)}")
print()

ok_all = True
for path in TARGETS:
    with io.open(path, encoding="utf-8", newline="") as f:
        text = f.read()

    fixed = text
    for fn, want in (("ПланБухгалтерии", PLAN_BUH), ("ПланКазны", PLAN_KAZNA)):
        pat = re.compile(
            r'(Функция\s+' + fn + r'\(\)\s+Экспорт\s*\n\s*Возврат\s+")([^"]*)(";)',
            re.UNICODE,
        )
        m = pat.search(fixed)
        if not m:
            print(f"[FAIL] {fn}: функция/литерал не найдены в {path}")
            ok_all = False
            continue

        got = m.group(2)
        if got == want:
            print(f"[OK]   {fn}() = {got}")
        else:
            print(f"[FIX]  {fn}(): в файле {got!r} ({cps(got)})")
            print(f"       из метаданных {want!r} ({cps(want)}) -> исправляю")
            fixed = pat.sub(lambda mm: mm.group(1) + want + mm.group(3), fixed, count=1)
            ok_all = False

    if fixed != text:
        with io.open(path, "w", encoding="utf-8", newline="") as f:
            f.write(fixed)
        print(f"[FIXED] файл перезаписан именами из метаданных: {path}")
        text = fixed

    n_kazna = text.count(PLAN_KAZNA)
    n_buh = text.count(PLAN_BUH)
    print(f"[CNT]  вхождений имени Казны в файле: {n_kazna} (требуется 1)")
    print(f"[CNT]  вхождений имени Бухгалтерии:   {n_buh} (требуется 1)")
    if n_kazna != 1 or n_buh != 1:
        ok_all = False

    # Регистр транспорта и маска каталога не должны использоваться в КОДЕ (ТЗ, критерий 1).
    # Упоминание в комментарии допустимо (объясняет, почему транспорт не используется).
    code_only = "\n".join(
        ln for ln in text.splitlines() if not ln.lstrip().startswith("//")
    )
    for forbidden in ("НастройкиТранспортаОбменаДанными", "НастройкиТранспортаОбмена",
                      "Индастриал", "СписокУзлов"):
        if forbidden in code_only:
            print(f"[FAIL] в КОДЕ остался запрещённый фрагмент: {forbidden}")
            ok_all = False
    print("[OK]   регистр транспорта и маска каталога в коде не используются")

print()
print("ИТОГ:", "OK — литералы совпадают с метаданными" if ok_all else "БЫЛИ ПРАВКИ/ОШИБКИ (см. выше)")
