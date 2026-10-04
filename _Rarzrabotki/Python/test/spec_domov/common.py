# -*- coding: utf-8 -*-
"""Общие функции П1 «Специфікація і еталон»: пути, COM bas_industrialbud, транзакция с откатом, файлы дампа.

База — ТОЛЬКО Srvr="localhost" (память feedback_kod_srvr_localhost). Данные в базу не пишутся:
всё, что тест создаёт, живёт внутри tx(B) и откатывается.
"""
import contextlib
import hashlib
import os
import sys

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

WT = r"C:\Configuration_downloads\BASERP25\.claude\worktrees\budget-calc-oct-dec-2026-20c998"
MAIN = r"C:\Configuration_downloads\BASERP25"
BB_WT = os.path.join(WT, r"_Rarzrabotki\BASEBuh")
BB_MAIN = os.path.join(MAIN, r"_Rarzrabotki\BASEBuh")
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
TMP = r"C:\Users\SUPPOR~1\AppData\Local\Temp\claude\plan_p1"
SKILLS = os.path.join(MAIN, r".claude\skills")
V8_BIN = r"C:\Program Files\1cv8\8.3.20.1914\bin"
CONN = 'Srvr="localhost";Ref="bas_industrialbud";Usr="cfo";Pwd="2442"'


def connect():
    import win32com.client
    return win32com.client.Dispatch("V83.COMConnector").Connect(CONN)


def err(e):
    """Текст ошибки 1С из pywintypes.com_error; иначе «Тип: текст» (AttributeError — опечатка в имени модуля/метода)."""
    info = getattr(e, "excepinfo", None)
    if info and info[2]:
        return str(info[2])
    return f"{type(e).__name__}: {e}"


def call(fn, *args):
    """(результат, None) или (None, текст ошибки 1С) — для негативных тестов.

    Ловит только com_error: AttributeError (опечатка в имени, функции нет в базе) летит дальше, иначе негативный
    тест «ошибка есть» прошёл бы впустую. Негативный тест проверяет подстроку текста ошибки, а не только её наличие.
    """
    import pywintypes
    try:
        return fn(*args), None
    except pywintypes.com_error as e:
        return None, err(e)


@contextlib.contextmanager
def tx(B):
    """Всё внутри — в транзакции, которая ВСЕГДА откатывается (вложенные НачатьТранзакцию из BSL тоже).

    Главная защита от записи в рабочую базу, поэтому громко падает (RuntimeError), если транзакцию не удержать:
    - на входе транзакция уже активна (вложенный tx или незакрытая чужая): внутренний откат закрыл бы все
      уровни, и код внешнего блока писал бы без транзакции — то есть фиксировал бы сразу;
    - к выходу транзакция уже неактивна: BSL закрыл её сверх своей НачатьТранзакцию — всё записанное после
      этого могло зафиксироваться. Исходная ошибка блока, если была, сохраняется в __cause__.
    Исключение BSL и ошибка записи транзакцию не закрывают (проверено COM 04.10.2026: после «не уникально»
    и после ВызватьИсключение TransactionActive() = True), так что негативные кейсы проверку не задевают.
    """
    if B.TransactionActive():
        raise RuntimeError("tx(): вложенный tx запрещён — транзакция уже активна; внутренний откат закрыл бы и внешнюю")
    B.BeginTransaction()
    failure = None
    try:
        yield
    except BaseException as e:
        failure = e
        raise
    finally:
        active = B.TransactionActive()
        while B.TransactionActive():
            B.RollbackTransaction()
        if not active:
            raise RuntimeError("tx(): транзакция закрыта внутри tx — данные могли записаться в базу") from failure


def arr(B, items):
    a = B.NewObject("Массив")
    for x in items:
        a.Добавить(x)
    return a


def q(B, text, **params):
    """Результат запроса списком словарей; значения — как отдаёт COM (ссылки — COM-объекты).

    Выгрузка по колонкам (ВыгрузитьКолонку): 3000 строк × 5 колонок — 2 с против 9 с при обращении к каждой ячейке.
    """
    z = B.NewObject("Запрос")
    z.Text = text
    for k, v in params.items():
        z.SetParameter(k, v)
    vt = z.Execute().Выгрузить()
    cols = [vt.Колонки.Получить(i).Имя for i in range(vt.Колонки.Количество())]
    n = vt.Количество()
    data = []
    for j in range(len(cols)):
        a = vt.ВыгрузитьКолонку(j)
        data.append([a.Получить(i) for i in range(n)])
    return [{c: data[j][i] for j, c in enumerate(cols)} for i in range(n)]


def vt_rows(vt, cols):
    """ТаблицаЗначений (COM) → список словарей по заданным колонкам."""
    return [{c: getattr(vt.Получить(i), c) for c in cols} for i in range(vt.Количество())]


def sv(B, v):
    """Строковое представление значения 1С (ссылки, перечисления — в языке сеанса, uk)."""
    if v is None or isinstance(v, (str, int, float, bool)):
        return v
    return B.String(v)


def same(B, a, b):
    """Равенство значений 1С (ссылок, перечислений) по ЗначениеВСтрокуВнутр: в строке и тип, и УИД.

    COM-обёртки сравнивать == нельзя; сравнение только по УИД склеивало пустые ссылки разных типов (все нули),
    по представлению — значения разных перечислений с одинаковым текстом.
    """
    if a is None or b is None:
        return a is b
    return B.ЗначениеВСтрокуВнутр(a) == B.ЗначениеВСтрокуВнутр(b)


def messages(B):
    m = B.ПолучитьСообщенияПользователю(True)
    return [str(m.Получить(i).Текст) for i in range(m.Количество())]


def changes_rows(B, ss_ref):
    """Строки регистрации СС к обмену: (узел, номер сообщения). Число строк при повторной регистрации не меняется — сравнивать сами строки."""
    rows = q(B, """ВЫБРАТЬ
	ПРЕДСТАВЛЕНИЕ(Изм.Узел) КАК Узел,
	ЕСТЬNULL(Изм.НомерСообщения, -1) КАК НомерСообщения
ИЗ
	Справочник.СтруктураСебестоимости.Изменения КАК Изм
ГДЕ
	Изм.Ссылка = &СС""", СС=ss_ref)
    return sorted((str(r["Узел"]), int(r["НомерСообщения"])) for r in rows)


def count_rows(B, table, where=""):
    return int(q(B, f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Н ИЗ {table} КАК Т {where}")[0]["Н"])


def read_text(path):
    """(текст с \\n, был_BOM, был_CRLF)."""
    with open(path, "rb") as f:
        raw = f.read()
    return raw.decode("utf-8-sig").replace("\r\n", "\n"), raw.startswith(b"\xef\xbb\xbf"), b"\r\n" in raw


def write_text(path, text, bom=True, crlf=True):
    """Канон дампа BASEBuh: UTF-8 с BOM и CRLF (память bsl_write_lf_crlf_canon)."""
    folder = os.path.dirname(path)
    if folder:
        os.makedirs(folder, exist_ok=True)
    text = text.replace("\r\n", "\n")
    if crlf:
        text = text.replace("\n", "\r\n")
    with open(path, "wb") as f:
        f.write((b"\xef\xbb\xbf" if bom else b"") + text.encode("utf-8"))


def md5_nocr(path):
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]
    return hashlib.md5(raw.replace(b"\r\n", b"\n")).hexdigest()


def zapros_path(name):
    return os.path.join(HERE, "zaprosy", name + ".sql")


def zapros(name):
    """Текст запроса — тот же файл, который build_bsl.py вставляет в BSL (Rule #-1: один источник)."""
    return read_text(zapros_path(name))[0].strip("\n")


def vidy_subkonto(B):
    """[Номенклатура, Склады] — склад не всегда третье субконто (E3 §3.2)."""
    pvh = B.ПланыВидовХарактеристик.ВидыСубконтоХозрасчетные
    return arr(B, [pvh.Номенклатура, pvh.Склады])
