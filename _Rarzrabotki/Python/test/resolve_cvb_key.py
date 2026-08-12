# -*- coding: utf-8 -*-
"""ЦВБ: расшифровка КлючСтроки в человекочитаемый вид (READ-ONLY).

Ключи движка — это UUID/коды через «|», в отчётах месяцев они печатаются «как есть».
Скрипт превращает их в имена объектов, чтобы решать: правило / исключение / реальный класс B.

Соответствие ключа контуру (Справочник.А_КонтурыСверкиБаз.КлючевыеПоля):
  Деньги безнал  = БанковскийСчетUID
  Товары         = НоменклатураКлюч|СкладКлюч            (оба = А_ИдКод = UUID справочника)
  Взаиморасчёты  = КонтрагентКлюч|ДоговорКлюч|ВалютаКод  (А_ИдКод контрагента/договора)
  Касса          = КассаUID

Запуск:
  python resolve_cvb_key.py Товары "a568a9cc-...|fe1c4b68-..." "57006114-...|b6b5907f-..."
  python resolve_cvb_key.py Касса e985b4c9-455b-11ed-80c0-00155d235309
"""
import sys

import win32com.client

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

if len(sys.argv) < 3:
    print(__doc__)
    sys.exit(2)

kontur = sys.argv[1]
keys = sys.argv[2:]

v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
S = erp.String


def by_uid(catalog_name, uid):
    """Ссылка справочника по UUID + человекочитаемое описание (или пометка «не найден»)."""
    try:
        ref = getattr(erp.Справочники, catalog_name).ПолучитьСсылку(
            erp.NewObject("УникальныйИдентификатор", uid))
        obj = ref.ПолучитьОбъект()
    except Exception as e:
        return f"<ошибка {catalog_name}: {e}>"
    if obj is None:
        return f"<{catalog_name}: объект по UID не найден>"
    text = str(S(obj.Наименование))
    for extra in ("ВидНоменклатуры", "ТипНоменклатуры", "Владелец", "Родитель"):
        if hasattr(obj, extra):
            try:
                val = str(S(getattr(obj, extra)))
                if val.strip():
                    text += f" | {extra}={val}"
            except Exception:
                pass
    return text


LAYOUT = {
    "Деньги безнал": [("БанковскийСчетUID", "БанковскиеСчетаОрганизаций")],
    "Товары": [("НоменклатураКлюч", "Номенклатура"), ("СкладКлюч", "Склады")],
    "Взаиморасчёты": [("КонтрагентКлюч", "Контрагенты"),
                      ("ДоговорКлюч", "ДоговорыКонтрагентов"),
                      ("ВалютаКод", None)],
    "Касса": [("КассаUID", "Кассы")],
}

if kontur not in LAYOUT:
    print(f"Неизвестный контур {kontur!r}. Известные: {list(LAYOUT)}")
    sys.exit(2)

for key in keys:
    print(f"\n=== [{kontur}] {key} ===")
    parts = key.split("|")
    for idx, (field, catalog) in enumerate(LAYOUT[kontur]):
        val = parts[idx] if idx < len(parts) else ""
        if not val:
            print(f"  {field} = <пусто>")
        elif catalog is None:
            print(f"  {field} = {val}")
        else:
            print(f"  {field} = {val}\n      -> {by_uid(catalog, val)}")

erp = None
print("\nRESOLVE DONE")
