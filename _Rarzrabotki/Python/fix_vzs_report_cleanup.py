# -*- coding: utf-8 -*-
"""
Чистка отчёта А_ВзаиморасчетыССотрудникамиДляБаланса.
Часть 1 (СКД, обе копии): убрать 4 поля периодов-корзин (День/Месяц/Квартал/Год)
                          из определений полей и из запроса; снять запятую после
                          ПериодСекунда; исправить укр. синоним.
Часть 2 (ObjectModule.bsl, только config Report): убрать процедуры #1-4
                          (ПриКомпоновкеРезультата, УстановитьОбязательныеНастройки,
                          СтруктураЗаголовковПолей, НастроитьПараметрыОтборыПоФункциональнымОпциям),
                          область ОбработчикиСобытий и вызов #4.
ПериодСекунда и блок дефолтного периода (#5-9) сохраняются.

Запуск: python fix_vzs_report_cleanup.py            (dry-run, ничего не пишет)
        python fix_vzs_report_cleanup.py --apply     (применить)
"""
import re
import sys
from pathlib import Path

ROOT = Path(r"C:\Configuration_downloads\BASERP25")

TEMPLATES = [
    ROOT / r"_Rarzrabotki\Отчеты\А_ВзаиморасчетыССотрудникамиДляБаланса\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml",
    ROOT / r"Reports\А_ВзаиморасчетыССотрудникамиДляБаланса\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml",
]
METADATA = [
    ROOT / r"_Rarzrabotki\Отчеты\А_ВзаиморасчетыССотрудникамиДляБаланса.xml",
    ROOT / r"Reports\А_ВзаиморасчетыССотрудникамиДляБаланса.xml",
]
OBJMODULES = [
    ROOT / r"Reports\А_ВзаиморасчетыССотрудникамиДляБаланса\Ext\ObjectModule.bsl",
    ROOT / r"_Rarzrabotki\Отчеты\А_ВзаиморасчетыССотрудникамиДляБаланса\Ext\ObjectModule.bsl",
]

PERIOD_BUCKETS = ["ПериодДень", "ПериодМесяц", "ПериодКвартал", "ПериодГод"]
SYN_OLD = "Відомість розрахунків з клієнтами (старий не використовувати)"
SYN_NEW = "Взаєморозрахунки зі співробітниками для балансу"


def read(path):
    raw = path.read_bytes()
    bom = raw.startswith(b"\xef\xbb\xbf")
    text = raw.decode("utf-8-sig")
    crlf = "\r\n" in text
    return text.replace("\r\n", "\n"), bom, crlf


def write(path, text, bom, crlf):
    out = text.replace("\n", "\r\n") if crlf else text
    data = (b"\xef\xbb\xbf" if bom else b"") + out.encode("utf-8")
    path.write_bytes(data)


def transform_template(t):
    # 1) удалить 4 блока <field> периодов-корзин (до </role></field>, чтобы не зацепить вложенный <field>)
    for name in PERIOD_BUCKETS:
        pat = (r'\n[ \t]*<field xsi:type="DataSetFieldField">\s*'
               r'<dataPath>' + name + r'</dataPath>.*?</role>\s*</field>')
        t = re.sub(pat, "", t, count=1, flags=re.S)
    # 2) удалить 4 строки выборки из запроса
    for name in PERIOD_BUCKETS:
        pat = r'\n[ \t]*А_ВзаиморасчетыССотрудникамиОстаткиИОбороты\.' + name + r' КАК ' + name + r',?'
        t = re.sub(pat, "", t, count=1)
    # 3) снять запятую после ПериодСекунда (теперь последнее поле перед ИЗ)
    t = t.replace(".ПериодСекунда КАК ПериодСекунда,", ".ПериодСекунда КАК ПериодСекунда")
    # ассерты результата
    for name in PERIOD_BUCKETS:
        assert name not in t, f"[Template] остался {name}"
    assert "ПериодСекунда" in t, "[Template] потеряна ПериодСекунда"
    assert "Регистратор" in t, "[Template] потерян Регистратор"
    assert ".ПериодСекунда КАК ПериодСекунда\n" in t or ".ПериодСекунда КАК ПериодСекунда " in t, \
        "[Template] запятая не снята корректно"
    return t


def transform_metadata(t):
    t = t.replace(SYN_OLD, SYN_NEW)
    assert SYN_OLD not in t, "[Metadata] старый синоним остался"
    assert SYN_NEW in t, "[Metadata] новый синоним отсутствует"
    return t


def transform_objmodule(t):
    # Идемпотентно: удаляем маркеры, корректность гарантируют финальные ассерты состояния
    # 1) убрать вызов #4 + комментарий в ПередЗагрузкойВариантаНаСервере
    pat_call = (r'\n[ \t]*// Изменение настроек по функциональным опциям'
                r'\n[ \t]*НастроитьПараметрыОтборыПоФункциональнымОпциям\([^\n]*\);')
    t = re.sub(pat_call, "", t, count=1)
    # 2) убрать всю область ОбработчикиСобытий (#1,#2,#3)
    pat_region = r'\n#Область ОбработчикиСобытий[ \t]*\n.*?\n#КонецОбласти\n'
    t = re.sub(pat_region, "", t, count=1, flags=re.S)
    # 3) убрать функцию #4 НастроитьПараметрыОтборыПоФункциональнымОпциям
    pat_fo = r'\nПроцедура НастроитьПараметрыОтборыПоФункциональнымОпциям\(.*?\nКонецПроцедуры\n'
    t = re.sub(pat_fo, "", t, count=1, flags=re.S)
    # 4) схлопнуть 2+ пустых/пробельных строк в одну
    t = re.sub(r'\n([ \t]*\n){2,}', "\n\n", t)
    # ассерты
    for gone in ["ПриКомпоновкеРезультата", "УстановитьОбязательныеНастройки",
                 "СтруктураЗаголовковПолей", "НастроитьПараметрыОтборыПоФункциональнымОпциям",
                 "#Область ОбработчикиСобытий", "КомпоновкаДанныхСервер", "ОтчетыСервер",
                 "КомпоновкаДанныхКлиентСервер"]:
        assert gone not in t, f"[ObjMod] остался: {gone}"
    for keep in ["ОпределитьНастройкиФормы", "ПередЗагрузкойВариантаНаСервере",
                 "НастроитьПараметрыОтборыПоУмолчанию", "ФиксированнаяНастройкаПараметра",
                 "ПользовательскаяНастройкаПараметра", "#Область СлужебныйПрограммныйИнтерфейс",
                 "#Область СлужебныеПроцедурыИФункции", "#КонецЕсли"]:
        assert keep in t, f"[ObjMod] потеряно: {keep}"
    assert t.count("#Область") == 2, f"[ObjMod] областей {t.count('#Область')} (ожидалось 2)"
    assert t.count("#КонецОбласти") == 2, f"[ObjMod] #КонецОбласти {t.count('#КонецОбласти')} (ожидалось 2)"
    return t


def process(path, fn, apply):
    text, bom, crlf = read(path)
    new = fn(text)
    changed = new != text
    d = len(text.splitlines()) - len(new.splitlines())
    print(f"  {'CHANGED' if changed else 'same   '}  -{d:>2} строк  BOM={int(bom)} CRLF={int(crlf)}  {path.name}  [{path.parent}]")
    if apply and changed:
        write(path, new, bom, crlf)
    return changed


def main():
    apply = "--apply" in sys.argv
    print(f"=== {'APPLY' if apply else 'DRY-RUN'} ===")
    print("Часть 1: СКД Template.xml")
    for p in TEMPLATES:
        process(p, transform_template, apply)
    print("Часть 1: синоним в метаданных")
    for p in METADATA:
        process(p, transform_metadata, apply)
    print("Часть 2: ObjectModule.bsl (обе копии)")
    for p in OBJMODULES:
        process(p, transform_objmodule, apply)
    print("OK — все ассерты пройдены.")


if __name__ == "__main__":
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    main()
