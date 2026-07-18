# -*- coding: utf-8 -*-
"""Пост-патч метаданных РасчетКомплектаций + СтатусыРасчетаКомплектаций:
uk-синонимы (§14), ChoiceFoldersAndItems=FoldersAndItems для счетов, DefaultObjectForm.
Обход последовательный по известному порядку узлов — имена колонок в разных ТЧ совпадают."""
import re, sys, io
import xml.dom.minidom

sys.stdout.reconfigure(encoding='utf-8')

CFG = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh"
DOC = CFG + r"\Documents\РасчетКомплектаций.xml"
ENUM = CFG + r"\Enums\СтатусыРасчетаКомплектаций.xml"

# Порядок узлов документа: (имя, uk, это_счет)
SEQ_DOC = [
    ("РасчетКомплектаций", "Розрахунок комплектацій за СС", False),
    # шапка
    ("Спецификация", "Специфікація (норми)", False),
    ("Период", "Залишки на дату", False),
    ("Организация", "Організація", False),
    ("ИсключатьЭтапы", "Виключати етапи зі списку", False),
    ("Статус", "Статус", False),
    ("Ответственный", "Відповідальний", False),
    ("Комментарий", "Коментар", False),
    # ТЧ СкладыОстатков
    ("СкладыОстатков", "Склади залишків", False),
    ("Склад", "Склад", False),
    # ТЧ СчетаОстатков
    ("СчетаОстатков", "Рахунки залишків", False),
    ("Счет", "Рахунок", True),
    # ТЧ СчетаМалоценки
    ("СчетаМалоценки", "Рахунки малоцінки", False),
    ("Счет", "Рахунок", True),
    # ТЧ Этапы
    ("Этапы", "Етапи", False),
    ("Этап", "Етап", False),
    # ТЧ ТабличнаяЧастьОстатков
    ("ТабличнаяЧастьОстатков", "Аналіз залишків", False),
    ("Номенклатура", "Номенклатура", False),
    ("ОбщееНазвание", "Загальна назва (еталон)", False),
    ("Склад", "Склад", False),
    ("Счет", "Рахунок", True),
    ("Остаток", "Залишок", False),
    ("Норма", "Норма (еталон)", False),
    ("ВНорме", "В нормі (по нормам)", False),
    ("ПонадНорму", "Понад норму (додаткова)", False),
    ("Экономия", "Економія", False),
    ("Единица", "Одиниця виміру залишків", False),
    ("ЕдиницаСС", "Одиниця виміру СС", False),
    ("СуммаОстатка", "Сума залишку", False),
    ("НормаСумма", "Сума норми", False),
    ("ВНормеСумма", "Сума в нормі", False),
    ("ПонадНормуСумма", "Сума понад", False),
    ("ЕкономіяСума", "Економія (грн)", False),
    ("НормаНоменкл", "Норма (номенкл.)", False),
    ("НормаСуммаНоменкл", "Сума норми (номенкл.)", False),
    ("Этап", "Етап", False),
    # ТЧ СписаниеПоНормам
    ("СписаниеПоНормам", "Списання за нормами", False),
    ("Номенклатура", "Номенклатура", False),
    ("ОбщееНазвание", "Загальна назва", False),
    ("Склад", "Склад", False),
    ("Счет", "Рахунок", True),
    ("Единица", "Одиниця", False),
    ("Количество", "Кількість", False),
    ("Цена", "Ціна (без ПДВ)", False),
    ("Сумма", "Сума (без ПДВ)", False),
    # ТЧ СписаниеСверхНормы
    ("СписаниеСверхНормы", "Списання понад норму", False),
    ("Номенклатура", "Номенклатура", False),
    ("ОбщееНазвание", "Загальна назва", False),
    ("Склад", "Склад", False),
    ("Счет", "Рахунок", True),
    ("Единица", "Одиниця", False),
    ("Количество", "Кількість", False),
    ("Цена", "Ціна (без ПДВ)", False),
    ("Сумма", "Сума (без ПДВ)", False),
    ("Причина", "Причина (обґрунтування)", False),
    # ТЧ ДокументиКомплектації
    ("ДокументиКомплектації", "Документи комплектації", False),
    ("Склад", "Склад", False),
    ("ДокументПоНормам", "Документ «По нормам»", False),
    ("ДокументДодаткова", "Документ «Понад норму»", False),
    # ТЧ ДокументиМалоценки
    ("ДокументиМалоценки", "Документи малоцінки", False),
    ("Склад", "Склад", False),
    ("ДокументПоНормам", "Документ «По нормам»", False),
    ("ДокументДодаткова", "Документ «Понад норму»", False),
    ("ФизЛицо", "МВО (фізична особа)", False),
    ("СпособОтраженияРасходов", "Спосіб відображення витрат", False),
]

SEQ_ENUM = [
    ("СтатусыРасчетаКомплектаций", "Статуси розрахунку комплектацій", False),
    ("Черновик", "Чернетка", False),
    ("РасчетВыполнен", "Розрахунок виконано", False),
    ("ДокументыСозданы", "Документи створено", False),
]


def read_text(path):
    raw = open(path, 'rb').read()
    bom = raw.startswith(b'\xef\xbb\xbf')
    return raw.decode('utf-8-sig'), bom


def write_text(path, text, bom):
    data = text.encode('utf-8')
    if bom:
        data = b'\xef\xbb\xbf' + data
    open(path, 'wb').write(data)


def patch_file(path, seq):
    text, bom = read_text(path)
    pos = 0
    patched_syn = 0
    patched_fold = 0
    for name, uk, is_account in seq:
        marker = f"<Name>{name}</Name>"
        i = text.find(marker, pos)
        assert i >= 0, f"не найден узел <Name>{name}</Name> после позиции {pos}"
        # конец блока Synonym этого узла
        syn_end = text.find("</Synonym>", i)
        assert syn_end >= 0, f"нет </Synonym> после {name}"
        # отступ: строка с </Synonym>
        line_start = text.rfind("\n", 0, syn_end) + 1
        indent = text[line_start:syn_end]  # табы перед </Synonym>
        uk_item = (f"{indent}\t<v8:item>\n"
                   f"{indent}\t\t<v8:lang>uk</v8:lang>\n"
                   f"{indent}\t\t<v8:content>{uk}</v8:content>\n"
                   f"{indent}\t</v8:item>\n")
        text = text[:line_start] + uk_item + text[line_start:]
        patched_syn += 1
        pos = syn_end + len(uk_item)
        if is_account:
            # первый ChoiceFoldersAndItems после этого узла (в пределах его Attribute-блока)
            j = text.find("<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>", pos)
            next_name = text.find("<Name>", pos)
            assert j >= 0 and (next_name < 0 or j < next_name), f"ChoiceFoldersAndItems не найден для {name} @{pos}"
            text = (text[:j] + "<ChoiceFoldersAndItems>FoldersAndItems</ChoiceFoldersAndItems>"
                    + text[j + len("<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>"):])
            patched_fold += 1
    return text, bom, patched_syn, patched_fold


# --- документ ---
text, bom, n_syn, n_fold = patch_file(DOC, SEQ_DOC)
# DefaultObjectForm
assert "<DefaultObjectForm/>" in text, "DefaultObjectForm уже заполнен?"
text = text.replace("<DefaultObjectForm/>",
                    "<DefaultObjectForm>Document.РасчетКомплектаций.Form.ФормаДокумента</DefaultObjectForm>", 1)
write_text(DOC, text, bom)
xml.dom.minidom.parse(DOC)
print(f"DOC OK: синонимов={n_syn}, FoldersAndItems={n_fold}, DefaultObjectForm задан")

# --- перечисление ---
text, bom, n_syn, n_fold = patch_file(ENUM, SEQ_ENUM)
write_text(ENUM, text, bom)
xml.dom.minidom.parse(ENUM)
print(f"ENUM OK: синонимов={n_syn}")
print("PATCH PASS")
