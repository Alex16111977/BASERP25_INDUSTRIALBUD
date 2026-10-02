# -*- coding: utf-8 -*-
"""Общие функции «мой расчётный бюджет ДДС производства, октябрь–декабрь 2026».

Подключение к BaseERP, даты через запрос (COM сдвигает даты на −3 ч), проекты производства
(проект = предок 2-го уровня СтруктураПредприятия под корнем «Производство» 0Ц-000004),
справочник статей ДДС, снимки данных в JSON (папка DATA) — расчётчик работает от снимков,
поэтому воспроизводим и не зависит от дневных перепроведений переноса Казны.
"""
import os, sys, json, datetime
from decimal import Decimal, ROUND_HALF_UP

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

CONN_ERP = 'Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"'
ROOT_CODE = "0Ц-000004"
BASE_DIR = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Бюджет прогноз\Бюджет октябрь 2026"
OUT_DIR = os.path.join(BASE_DIR, "Мой расчет")
DATA = os.path.join(OUT_DIR, "data")
os.makedirs(DATA, exist_ok=True)

# Транзит и финансирование — не операционные деньги проекта (комиссия финагента УТ-001097 — расход).
ТРАНЗИТ = {"УТ-001210", "УТ-001211"}
ФИНАНСИРОВАНИЕ = {"УТ-001201", "УТ-001202", "УТ-001082", "УТ-001096"}

_erp = None


def conn():
    global _erp
    if _erp is None:
        import win32com.client
        v8 = win32com.client.Dispatch("V83.COMConnector")
        _erp = v8.Connect(CONN_ERP)
    return _erp


def S(x):
    return conn().String(x)


def uid(ref):
    """УИД ссылки строкой (COM-прокси одной ссылки в Python не равны через ==)."""
    return conn().string(ref.УникальныйИдентификатор())


def dt(y, m, d, h=0, mi=0, s=0):
    """Дата на стороне 1С без сдвига часового пояса COM."""
    q = conn().NewObject("Запрос")
    q.Text = f"ВЫБРАТЬ ДАТАВРЕМЯ({y},{m},{d},{h},{mi},{s}) КАК Д"
    return q.Execute().Выгрузить().Получить(0).Д


def query(text, **params):
    q = conn().NewObject("Запрос")
    q.Text = text
    for k, v in params.items():
        if isinstance(v, (list, tuple, set)):
            arr = conn().NewObject("Массив")
            for x in v:
                arr.Добавить(x)
            v = arr
        q.SetParameter(k, v)
    try:
        return q.Execute().Выгрузить()
    except Exception as e:
        info = getattr(e, "excepinfo", None)
        raise RuntimeError(f"запрос упал: {info[2] if info else e}\n{text}")


def ym(date_1c):
    """Дата 1С -> 'YYYY-MM' (берём строку самой 1С: dd.mm.yyyy)."""
    s = S(date_1c)
    return f"{s[6:10]}-{s[3:5]}"


def ymd(date_1c):
    s = S(date_1c)
    return f"{s[6:10]}-{s[3:5]}-{s[0:2]}"


def окр(x):
    """Как типизированная колонка 1С (Число 15,2): арифметическое округление."""
    return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def fmt(x, d=2):
    return f"{float(x):,.{d}f}".replace(",", " ").replace(".", ",")


def save(name, obj):
    path = os.path.join(DATA, name)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1, default=str)
    print(f"  снимок: {path}")
    return path


def load(name):
    with open(os.path.join(DATA, name), encoding="utf-8") as f:
        return json.load(f)


def snapshot_stamp():
    return datetime.datetime.now().strftime("%d.%m.%Y %H:%M")


def projects():
    """Подразделение (УИД) -> проект {uid, name, code}; плюс УИД корня (корень = ЦО «Производство»)."""
    t = query("""ВЫБРАТЬ П.Ссылка КАК Ссылка ИЗ Справочник.СтруктураПредприятия КАК П ГДЕ П.Код = &Код""", Код=ROOT_CODE)
    root = t.Получить(0).Ссылка
    t = query("""ВЫБРАТЬ
        СП.Ссылка КАК П0, СП.Наименование КАК Имя, СП.Код КАК Код,
        СП.Родитель КАК П1, СП.Родитель.Родитель КАК П2, СП.Родитель.Родитель.Родитель КАК П3,
        СП.Родитель.Родитель.Родитель.Родитель КАК П4, СП.Родитель.Родитель.Родитель.Родитель.Родитель КАК П5
    ИЗ Справочник.СтруктураПредприятия КАК СП
    ГДЕ СП.Ссылка В ИЕРАРХИИ (&Корень)""", Корень=root)
    erp = conn()
    names, codes, chains = {}, {}, {}
    for i in range(t.Количество()):
        r = t.Получить(i)
        u = uid(r.П0)
        names[u] = S(r.Имя).strip()
        codes[u] = S(r.Код).strip()
        chain = []
        for c in ("П0", "П1", "П2", "П3", "П4", "П5"):
            v = getattr(r, c)
            if not erp.ЗначениеЗаполнено(v):
                break
            chain.append(uid(v))
        chains[u] = chain
    root_u = uid(root)
    m = {}
    for u, chain in chains.items():
        # проект — элемент цепи сразу под корнем «Производство»; сам корень — ЦО
        k = chain.index(root_u) if root_u in chain else len(chain) - 1
        pu = chain[k - 1] if k >= 1 else root_u
        m[u] = {"uid": pu, "name": names.get(pu, "?"), "code": codes.get(pu, "?"), "dept": names[u]}
    return m, root_u


def articles():
    """Статьи ДДС: УИД -> {code, name, type(Приход/Расход/''), parent, forbidden(ветка УТ-000965)}."""
    t = query("""ВЫБРАТЬ
        Ст.Ссылка КАК Ссылка, Ст.Код КАК Код, Ст.Наименование КАК Имя, Ст.Родитель.Наименование КАК Родитель,
        ВЫБОР КОГДА Ст.А_ТипПоказателя = ЗНАЧЕНИЕ(Перечисление.ТипыЗначенийПоказателейБюджетногоОтчета.Приход) ТОГДА "Приход"
             КОГДА Ст.А_ТипПоказателя = ЗНАЧЕНИЕ(Перечисление.ТипыЗначенийПоказателейБюджетногоОтчета.Расход) ТОГДА "Расход"
             ИНАЧЕ "" КОНЕЦ КАК Тип,
        Ст.Ссылка В ИЕРАРХИИ (ВЫБРАТЬ Г.Ссылка ИЗ Справочник.СтатьиДвиженияДенежныхСредств КАК Г ГДЕ Г.Код = "УТ-000965") КАК Запрет,
        Ст.ЭтоГруппа КАК Группа, Ст.ПометкаУдаления КАК Пометка
    ИЗ Справочник.СтатьиДвиженияДенежныхСредств КАК Ст""")
    out = {}
    for i in range(t.Количество()):
        r = t.Получить(i)
        out[uid(r.Ссылка)] = {"code": S(r.Код).strip(), "name": S(r.Имя).strip(), "type": S(r.Тип),
                             "parent": S(r.Родитель).strip(), "forbidden": bool(r.Запрет),
                             "group": bool(r.Группа), "deleted": bool(r.Пометка)}
    return out


def article_by_code(code):
    """Ссылка на статью ДДС по коду: не группа, не помечена, вне запретной ветки УТ-000965."""
    t = query("""ВЫБРАТЬ СТ.Ссылка КАК Ссылка ИЗ Справочник.СтатьиДвиженияДенежныхСредств КАК СТ
        ГДЕ СТ.Код = &Код И НЕ СТ.ЭтоГруппа И НЕ СТ.ПометкаУдаления
          И НЕ СТ.Ссылка В ИЕРАРХИИ (ВЫБРАТЬ Г.Ссылка ИЗ Справочник.СтатьиДвиженияДенежныхСредств КАК Г ГДЕ Г.Код = "УТ-000965")""",
              Код=code)
    assert t.Количество() == 1, f"статья {code}: найдено {t.Количество()}"
    return t.Получить(0).Ссылка
