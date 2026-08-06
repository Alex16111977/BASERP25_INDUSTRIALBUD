# -*- coding: utf-8 -*-
"""Строит Templates/МакетАнализЦех из МакетАнализССОдна (печать для кладовщика).

Колонки: Норма СС | Видано | Одиниця | В нормі | Понад норму | Економія | Коментар цеху.
Норма/Економія — только на строке группы. Жирные группы, жёлтым только понаднорма.
Идемпотентен: перезапись целевых файлов, регистрация в корневом xml — однократная.
"""
import os
import re
import sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

# путь от расположения скрипта: работает и в worktree, и в основной конфигурации
BASE = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                     r"..\..\BASEBuh\Documents\РасчетКомплектаций"))
SRC_TPL = BASE + r"\Templates\МакетАнализССОдна\Ext\Template.xml"
SRC_WRAP = BASE + r"\Templates\МакетАнализССОдна.xml"
DST_DIR = BASE + r"\Templates\МакетАнализЦех"
DST_TPL = DST_DIR + r"\Ext\Template.xml"
DST_WRAP = BASE + r"\Templates\МакетАнализЦех.xml"
ROOT_XML = BASE + r"\..\РасчетКомплектаций.xml"
UUID_NEW = "64dd02bb-7ffb-4d3d-b256-4a4b0eb21246"

t = open(SRC_TPL, encoding="utf-8-sig").read()

def cut_block(text, open_tag, close_tag, marker):
    """Вырезает первый блок open_tag...close_tag, содержащий marker."""
    start = 0
    while True:
        i = text.find(open_tag, start)
        assert i >= 0, f"нет блока с {marker!r}"
        j = text.find(close_tag, i) + len(close_tag)
        if marker in text[i:j]:
            k = text.rfind("\n", 0, i)
            return text[:k] + text[j:], text[i:j]
        start = j

def cut_ca_item(text, marker):
    """Вырезает элемент <dcsset:item>…</dcsset:item> условного оформления, содержащий
    marker, с учётом ВЛОЖЕННЫХ dcsset:item (FilterItemComparison и т.п.)."""
    m = text.find(marker)
    assert m >= 0, f"нет маркера {marker!r}"
    i = text.rfind("<dcsset:item>", 0, m)
    assert i >= 0, f"нет открытия <dcsset:item> перед {marker!r}"
    pos, depth = i, 0
    while True:
        no = text.find("<dcsset:item", pos)
        nc = text.find("</dcsset:item>", pos)
        assert nc >= 0, "разбалансированы dcsset:item"
        if 0 <= no < nc:
            depth += 1
            pos = no + len("<dcsset:item")
        else:
            depth -= 1
            pos = nc + len("</dcsset:item>")
            if depth == 0:
                break
    k = text.rfind("\n", 0, i)
    return text[:k] + text[pos:]

def get_block(text, open_tag, close_tag, marker):
    start = 0
    while True:
        i = text.find(open_tag, start)
        assert i >= 0, f"нет блока с {marker!r}"
        j = text.find(close_tag, i) + len(close_tag)
        if marker in text[i:j]:
            return text[i:j]
        start = j

# --- 1. Удалить ненужные поля набора данных и вычисляемые поля ---
# закрывающий якорь с отступом: внутри блока есть однострочный <field>Имя</field>
FIELD_CLOSE = "\n\t\t</field>"
for fld in ("НормаНоменкл", "ВНормеСумма", "ПонадНормуСумма", "ЕкономіяСума",
            "НормаСуммаНоменкл", "СуммаОстатка"):
    t, _ = cut_block(t, "<field xsi:type=", FIELD_CLOSE, f"<dataPath>{fld}</dataPath>")
t, _ = cut_block(t, "<calculatedField>", "</calculatedField>", "<dataPath>РизницяСС</dataPath>")
t, _ = cut_block(t, "<calculatedField>", "</calculatedField>", "<dataPath>ПроцВиконанняСС</dataPath>")

# --- 2. Добавить поле Норма (клон блока ВНорме) с титулом «Норма СС» ---
vnorme = get_block(t, "<field xsi:type=", FIELD_CLOSE, "<dataPath>ВНорме</dataPath>")
norma = vnorme.replace("<dataPath>ВНорме</dataPath>", "<dataPath>Норма</dataPath>")
norma = norma.replace("<field>ВНорме</field>", "<field>Норма</field>")
norma = re.sub(r"<v8:content>[^<]*</v8:content>", "<v8:content>Норма СС</v8:content>", norma)
i = t.find("<field xsi:type=")  # вставить перед первым полем
k = t.rfind("\n", 0, i)
indent = t[k + 1:i]
t = t[:i] + norma + "\n" + indent + t[i:]

# --- 3. Добавить вычисляемое поле КоментарЦеху (пустая строка) ---
comment_field = (
    "<calculatedField>\n"
    "\t\t<dataPath>КоментарЦеху</dataPath>\n"
    "\t\t<expression>\"\"</expression>\n"
    "\t\t<title xsi:type=\"v8:LocalStringType\">\n"
    "\t\t\t<v8:item>\n"
    "\t\t\t\t<v8:lang>uk</v8:lang>\n"
    "\t\t\t\t<v8:content>Коментар цеху</v8:content>\n"
    "\t\t\t</v8:item>\n"
    "\t\t</title>\n"
    "\t</calculatedField>")
anchor = "<totalField>"
i = t.find(anchor)
assert i > 0
k = t.rfind("\n", 0, i)
indent = t[k + 1:i]
t = t[:i] + comment_field + "\n" + indent + t[i:]

# --- 4. totalFields: удалить лишние (включая агрегаты РизницяСС/ПроцВиконанняСС), добавить Норма ---
for fld in ("НормаНоменкл", "ВНормеСумма", "ПонадНормуСумма", "ЕкономіяСума",
            "НормаСуммаНоменкл", "СуммаОстатка", "РизницяСС", "ПроцВиконанняСС"):
    t, _ = cut_block(t, "<totalField>", "</totalField>", f"<dataPath>{fld}</dataPath>")
tf_vnorme = get_block(t, "<totalField>", "</totalField>", "<dataPath>ВНорме</dataPath>")
tf_norma = tf_vnorme.replace("ВНорме", "Норма")
# КоментарЦеху — тоже ресурс (Максимум над ""): иначе Auto-уровни его не включают,
# а не-ресурсные поля СКД выводит ПЕРЕД ресурсами (колонка прыгает в начало)
tf_comment = ("<totalField>\n"
              "\t\t<dataPath>КоментарЦеху</dataPath>\n"
              "\t\t<expression>Максимум(КоментарЦеху)</expression>\n"
              "\t</totalField>")
i = t.find("<totalField>")
k = t.rfind("\n", 0, i)
indent = t[k + 1:i]
t = t[:i] + tf_norma + "\n" + indent + tf_comment + "\n" + indent + t[i:]

# --- 5. Глобальный selection: 7 полей в порядке колонок ---
def selection_xml(fields, indent):
    items = "".join(
        f"{indent}\t<dcsset:item xsi:type=\"dcsset:SelectedItemField\">\n"
        f"{indent}\t\t<dcsset:field>{f}</dcsset:field>\n"
        f"{indent}\t</dcsset:item>\n" for f in fields)
    return f"<dcsset:selection>\n{items}{indent}</dcsset:selection>"

ALL7 = ["Норма", "Остаток", "ЕдиницаЕдина", "ВНорме", "ПонадНорму", "Экономия", "КоментарЦеху"]
DET5 = ["Остаток", "ЕдиницаЕдина", "ВНорме", "ПонадНорму", "КоментарЦеху"]

i = t.find("<dcsset:selection>")
j = t.find("</dcsset:selection>", i) + len("</dcsset:selection>")
t = t[:i] + selection_xml(ALL7, "\t\t\t") + t[j:]

# --- 6. Титул Остаток -> Видано ---
ost = get_block(t, "<field xsi:type=", FIELD_CLOSE, "<dataPath>Остаток</dataPath>")
t = t.replace(ost, re.sub(r"<v8:content>[^<]*</v8:content>",
                          "<v8:content>Видано</v8:content>", ost), 1)

# --- 7. Условное оформление: удалить зелёное, янтарное перевести на ПонадНорму ---
t = cut_ca_item(t, "#E2EFDA")
t = t.replace("<dcsset:left xsi:type=\"dcscor:Field\">ПонадНормуСумма</dcsset:left>",
              "<dcsset:left xsi:type=\"dcscor:Field\">ПонадНорму</dcsset:left>")

# --- 8. Заголовок и представление варианта ---
t = t.replace("<v8:content>Аналіз СС (одна одиниця)</v8:content>",
              "<v8:content>Аналіз залишків для списання за СС</v8:content>")

# --- 9. Уровни структуры: Auto-selection ОСТАВЛЯЕМ (единая колонка-дерево названий);
#         группе — жирный шрифт; деталям — глушим Норма/Экономия параметром «Текст» = "" ---
auto_sel = ("<dcsset:selection>\n\t\t\t\t\t<dcsset:item xsi:type=\"dcsset:SelectedItemAuto\"/>"
            "\n\t\t\t\t</dcsset:selection>")
assert t.count(auto_sel) == 1, f"ожидался 1 Auto-selection группы, найдено {t.count(auto_sel)}"
bold_ca = (
    "<dcsset:conditionalAppearance>\n"
    "\t\t\t\t\t<dcsset:item>\n"
    "\t\t\t\t\t\t<dcsset:selection/>\n"
    "\t\t\t\t\t\t<dcsset:appearance>\n"
    "\t\t\t\t\t\t\t<dcscor:item xsi:type=\"dcsset:SettingsParameterValue\">\n"
    "\t\t\t\t\t\t\t\t<dcscor:parameter>Шрифт</dcscor:parameter>\n"
    "\t\t\t\t\t\t\t\t<dcscor:value xsi:type=\"v8ui:Font\" ref=\"sys:DefaultGUIFont\" bold=\"true\" italic=\"false\" underline=\"false\" strikeout=\"false\" kind=\"WindowsFont\"/>\n"
    "\t\t\t\t\t\t\t</dcscor:item>\n"
    "\t\t\t\t\t\t</dcsset:appearance>\n"
    "\t\t\t\t\t\t<dcsset:useInFieldsHeader>DontUse</dcsset:useInFieldsHeader>\n"
    "\t\t\t\t\t\t<dcsset:useInHeader>DontUse</dcsset:useInHeader>\n"
    "\t\t\t\t\t\t<dcsset:useInParameters>DontUse</dcsset:useInParameters>\n"
    "\t\t\t\t\t</dcsset:item>\n"
    "\t\t\t\t</dcsset:conditionalAppearance>")
t = t.replace(auto_sel, auto_sel + "\n\t\t\t\t" + bold_ca, 1)

def mute_ca(fields, indent):
    """СА детального уровня: пустой текст для полей fields (значения только на группе)."""
    sel = "".join(
        f"{indent}\t\t\t<dcsset:item>\n"
        f"{indent}\t\t\t\t<dcsset:field>{f}</dcsset:field>\n"
        f"{indent}\t\t\t</dcsset:item>\n" for f in fields)
    return (
        f"<dcsset:conditionalAppearance>\n"
        f"{indent}\t<dcsset:item>\n"
        f"{indent}\t\t<dcsset:selection>\n{sel}"
        f"{indent}\t\t</dcsset:selection>\n"
        f"{indent}\t\t<dcsset:appearance>\n"
        f"{indent}\t\t\t<dcscor:item xsi:type=\"dcsset:SettingsParameterValue\">\n"
        f"{indent}\t\t\t\t<dcscor:parameter>Текст</dcscor:parameter>\n"
        f"{indent}\t\t\t\t<dcscor:value xsi:type=\"xs:string\"></dcscor:value>\n"
        f"{indent}\t\t\t</dcscor:item>\n"
        f"{indent}\t\t</dcsset:appearance>\n"
        f"{indent}\t\t<dcsset:useInFieldsHeader>DontUse</dcsset:useInFieldsHeader>\n"
        f"{indent}\t\t<dcsset:useInHeader>DontUse</dcsset:useInHeader>\n"
        f"{indent}\t\t<dcsset:useInParameters>DontUse</dcsset:useInParameters>\n"
        f"{indent}\t</dcsset:item>\n"
        f"{indent}</dcsset:conditionalAppearance>")

auto_sel_det = ("<dcsset:selection>\n\t\t\t\t\t\t<dcsset:item xsi:type=\"dcsset:SelectedItemAuto\"/>"
                "\n\t\t\t\t\t</dcsset:selection>")
assert t.count(auto_sel_det) == 1, f"ожидался 1 Auto-selection деталей, найдено {t.count(auto_sel_det)}"
t = t.replace(auto_sel_det,
              auto_sel_det + "\n\t\t\t\t\t" + mute_ca(["Норма", "Экономия"], "\t\t\t\t\t"), 1)

# --- 10. Записать Template.xml ---
os.makedirs(DST_DIR + r"\Ext", exist_ok=True)
with open(DST_TPL, "w", encoding="utf-8-sig", newline="\n") as f:
    f.write(t)

# --- 11. Обёртка с фиксированным uuid ---
w = open(SRC_WRAP, encoding="utf-8-sig").read()
w = re.sub(r'uuid="[0-9a-f-]+"', f'uuid="{UUID_NEW}"', w)
w = w.replace("<Name>МакетАнализССОдна</Name>", "<Name>МакетАнализЦех</Name>")
w = w.replace("<v8:content>Макет аналіз СС (одна одиниця)</v8:content>",
              "<v8:content>Макет аналіз для цеху</v8:content>")
with open(DST_WRAP, "w", encoding="utf-8-sig", newline="\n") as f:
    f.write(w)

# --- 12. Регистрация в корневом xml документа ---
root_path = os.path.normpath(ROOT_XML)
r = open(root_path, encoding="utf-8-sig").read()
if "<Template>МакетАнализЦех</Template>" not in r:
    r = r.replace("<Template>МакетАнализССОдна</Template>",
                  "<Template>МакетАнализССОдна</Template>\n\t\t\t<Template>МакетАнализЦех</Template>")
    with open(root_path, "w", encoding="utf-8-sig", newline="\n") as f:
        f.write(r)
    print("Корневой xml: макет зарегистрирован")
else:
    print("Корневой xml: регистрация уже есть")

# --- 13. Самопроверки ---
import xml.etree.ElementTree as ET
ET.fromstring(open(DST_TPL, encoding="utf-8-sig").read().encode("utf-8"))
chk = open(DST_TPL, encoding="utf-8-sig").read()
for absent in ("НормаНоменкл", "ВНормеСумма", "ПонадНормуСумма", "ЕкономіяСума",
               "НормаСуммаНоменкл", "СуммаОстатка", "РизницяСС", "ПроцВиконанняСС", "#E2EFDA"):
    assert absent not in chk, f"осталось: {absent}"
for present in ("<dataPath>Норма</dataPath>", "КоментарЦеху", "Видано", "Норма СС",
                "#FFE699", 'bold="true"',
                "<dcsset:left xsi:type=\"dcscor:Field\">ПонадНорму</dcsset:left>"):
    assert present in chk, f"нет: {present}"
assert chk.count("SelectedItemAuto") == 2, "уровневые Auto-selection должны остаться (2)"
assert "<dcscor:parameter>Текст</dcscor:parameter>" in chk, "нет глушилки Норма/Экономия у деталей"
print("OK: МакетАнализЦех построен и проверен")
