# -*- coding: utf-8 -*-
"""Самопроверка common.py (только чтение базы): транзакция всегда откатывается, ошибка 1С читается, BOM+CRLF, md5 без CR."""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common  # noqa: E402
from common import call, connect, md5_nocr, messages, q, read_text, same, tx, write_text, zapros_path  # noqa: E402


def expect_runtime(needle, body):
    """body() должна упасть RuntimeError с needle в тексте; возвращает исключение."""
    try:
        body()
    except RuntimeError as e:
        assert needle in str(e), e
        return e
    raise AssertionError("нет RuntimeError «" + needle + "»")


def test_tx_rollback(B):
    with tx(B):
        assert B.TransactionActive()
        B.BeginTransaction()                     # как будто BSL открыл вложенную и упал
    assert not B.TransactionActive()


def test_tx_nested_forbidden(B):
    """Вложенный tx: внутренний откат закрыл бы все уровни, и внешний блок писал бы без транзакции."""
    def body():
        with tx(B):
            with tx(B):
                raise AssertionError("вложенный tx пропущен")
    expect_runtime("вложенный tx запрещён", body)
    assert not B.TransactionActive()


def test_tx_closed_inside(B):
    """BSL закрыл транзакцию сверх своей НачатьТранзакцию — всё дальнейшее фиксировалось бы сразу."""
    def body():
        with tx(B):
            B.RollbackTransaction()
    expect_runtime("закрыта внутри tx", body)
    assert not B.TransactionActive()

    def body_err():
        with tx(B):
            B.RollbackTransaction()
            raise ValueError("исходная ошибка теста")
    e = expect_runtime("закрыта внутри tx", body_err)
    assert isinstance(e.__cause__, ValueError), "исходная ошибка потеряна"
    assert not B.TransactionActive()


def test_tx_bsl_error_keeps_tx(B):
    """Исключение BSL внутри tx транзакцию не закрывает: негативные кейсы (call, expect_write_error) не ловят ложную тревогу."""
    with tx(B):
        res, text = call(B.ОбщегоНазначенияКлиентСервер.ПроверитьПараметр, "ПробаTx", "П", 1,
                         B.NewObject("ОписаниеТипов", "Строка"))
        assert res is None and "ПробаTx" in text, text
        assert B.TransactionActive()
    assert not B.TransactionActive()


def test_err_and_q(B):
    rows = q(B, "ВЫБРАТЬ 1 КАК Н")
    assert rows == [{"Н": 1}], rows
    rows = q(B, """ВЫБРАТЬ 2 КАК А, "y" КАК Б
ОБЪЕДИНИТЬ ВСЕ
ВЫБРАТЬ 1, "x"
УПОРЯДОЧИТЬ ПО А""")
    assert rows == [{"А": 1, "Б": "x"}, {"А": 2, "Б": "y"}], rows
    assert q(B, "ВЫБРАТЬ Т.Ссылка КАК С ИЗ Справочник.Валюты КАК Т ГДЕ ЛОЖЬ") == []
    z = B.NewObject("Запрос")
    z.Text = "ВЫБРАТЬ ОШИБКА ИЗ"
    res, text = call(z.Execute)
    assert res is None and text, text
    assert isinstance(messages(B), list)
    try:                                         # опечатка в Python-коде не превращается в «ожидаемую ошибку 1С»
        call(lambda: B.НетТакогоИмени_П1)
    except AttributeError:
        pass
    else:
        raise AssertionError("call() проглотил AttributeError")


def test_same(B):
    nom, skl = B.Справочники.Номенклатура.ПустаяСсылка(), B.Справочники.Склады.ПустаяСсылка()
    assert not same(B, nom, skl), "пустые ссылки разных типов равны"
    assert same(B, nom, B.Справочники.Номенклатура.ПустаяСсылка())
    text = """ВЫБРАТЬ ПЕРВЫЕ 2 Т.Ссылка КАК С ИЗ Справочник.Номенклатура КАК Т УПОРЯДОЧИТЬ ПО Т.Код"""
    a, b = [r["С"] for r in q(B, text)]
    assert same(B, a, q(B, text)[0]["С"]) and not same(B, a, b)
    assert same(B, None, None) and not same(B, None, a)


def test_files():
    with tempfile.TemporaryDirectory() as d:
        p1, p2 = os.path.join(d, "a.bsl"), os.path.join(d, "b.bsl")
        write_text(p1, "Рядок 1\nРядок 2\n")
        with open(p1, "rb") as f:
            raw = f.read()
        assert raw.startswith(b"\xef\xbb\xbf"), "нет BOM"
        assert raw.count(b"\r\n") == 2 and raw.replace(b"\r\n", b"").count(b"\n") == 0, "не CRLF"
        with open(p2, "wb") as f:
            f.write("Рядок 1\nРядок 2\n".encode("utf-8"))
        assert md5_nocr(p1) == md5_nocr(p2)
        text, bom, crlf = read_text(p1)
        assert text == "Рядок 1\nРядок 2\n" and bom and crlf
        assert zapros_path("x").endswith(os.path.join("zaprosy", "x.sql"))

        p_sql = os.path.join(d, "x.sql")
        write_text(p_sql, "\nВЫБРАТЬ 1\nКАК Н\n\n")
        orig = common.zapros_path
        common.zapros_path = lambda name: p_sql
        try:
            assert common.zapros("x") == "ВЫБРАТЬ 1\nКАК Н", repr(common.zapros("x"))
        finally:
            common.zapros_path = orig

        cwd = os.getcwd()
        os.chdir(d)
        try:
            write_text("c.txt", "x\n")           # путь без каталога
        finally:
            os.chdir(cwd)
        assert read_text(os.path.join(d, "c.txt")) == ("x\n", True, True)


if __name__ == "__main__":
    B = connect()
    test_tx_rollback(B)
    test_tx_nested_forbidden(B)
    test_tx_closed_inside(B)
    test_tx_bsl_error_keeps_tx(B)
    test_err_and_q(B)
    test_same(B)
    test_files()
    print("OK test_common")
