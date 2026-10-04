# -*- coding: utf-8 -*-
"""postpatch.fix: cfg: вместо d5p1:, '/>' вместо ' />', uk к каждому ru, идемпотентность, отказ без перевода."""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import postpatch  # noqa: E402
from common import read_text, write_text  # noqa: E402

HEAD = ('<?xml version="1.0" encoding="UTF-8"?>\n<MetaDataObject xmlns="http://v8.1c.ru/8.3/MDClasses" '
        'xmlns:app="http://v8.1c.ru/8.2/managed-application/core" xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" '
        'xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.13">\n')


def syn(ru, t):
    return (f"{t}<Synonym>\n{t}\t<v8:item>\n{t}\t\t<v8:lang>ru</v8:lang>\n{t}\t\t<v8:content>{ru}</v8:content>\n"
            f"{t}\t</v8:item>\n{t}</Synonym>\n")


def attr(name, ru):
    return (f'\t\t\t<Attribute uuid="{name}-u">\n\t\t\t\t<Properties>\n\t\t\t\t\t<Name>{name}</Name>\n' + syn(ru, "\t\t\t\t\t")
            + '\t\t\t\t\t<Type>\n\t\t\t\t\t\t<v8:Type xmlns:d5p1="http://v8.1c.ru/8.1/data/enterprise/current-config">'
              'd5p1:CatalogRef.А_СпецификацииДомов</v8:Type>\n\t\t\t\t\t</Type>\n\t\t\t\t\t<ChoiceParameters />\n'
              '\t\t\t\t</Properties>\n\t\t\t</Attribute>\n')


def test_enum_like():
    p = os.path.join(tempfile.mkdtemp(), "e.xml")
    write_text(p, HEAD + '\t<Enum uuid="1">\n\t\t<Properties>\n\t\t\t<Name>А_Тест</Name>\n' + syn("Узел", "\t\t\t")
               + "\t\t\t<Comment />\n\t\t</Properties>\n\t</Enum>\n</MetaDataObject>\n")
    postpatch.fix(p)
    text, bom, crlf = read_text(p)
    assert bom and crlf and " />" not in text
    assert "<v8:lang>uk</v8:lang>\n\t\t\t\t\t<v8:content>Вузол</v8:content>" in text, text
    postpatch.fix(p)
    assert read_text(p)[0].count("<v8:lang>uk</v8:lang>") == 1, "не идемпотентен"


def test_catalog_choice_and_cfg():
    p = os.path.join(tempfile.mkdtemp(), "c.xml")
    body = attr("Узел", "Узел") + attr("Основание", "Создано из") + attr("Узел", "Узел") + attr("Узел", "Узел строителя")
    write_text(p, HEAD + '\t<Catalog uuid="2">\n\t\t<Properties>\n\t\t\t<CheckUnique>false</CheckUnique>\n'
               "\t\t</Properties>\n\t\t<ChildObjects>\n" + body + "\t\t</ChildObjects>\n\t</Catalog>\n</MetaDataObject>\n")
    postpatch.fix(p, catalog=True)
    text = read_text(p)[0]
    assert "d5p1" not in text and text.count("<v8:Type>cfg:CatalogRef.А_СпецификацииДомов</v8:Type>") == 4
    assert text.count("Enum.А_ВидыСпецификацийДомов.EnumValue.Узел") == 3
    assert "<CheckUnique>true</CheckUnique>" in text


def test_missing_translation():
    p = os.path.join(tempfile.mkdtemp(), "m.xml")
    write_text(p, HEAD + '\t<Enum uuid="1">\n\t\t<Properties>\n' + syn("Слово без перевода", "\t\t\t")
               + "\t\t</Properties>\n\t</Enum>\n</MetaDataObject>\n")
    try:
        postpatch.fix(p)
    except AssertionError as e:
        assert "Слово без перевода" in str(e)
        return
    raise AssertionError("fix должен отказать без uk-перевода")


if __name__ == "__main__":
    test_enum_like()
    test_catalog_choice_and_cfg()
    test_missing_translation()
    print("OK test_postpatch")
