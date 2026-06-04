# -*- coding: utf-8 -*-
"""
Общий модуль для скриптов knowledge_Balanse_klient.

ЧТО ДЕЛАЕТ:
- Подключается к ERP через V83.COMConnector.
- Возвращает готовые ссылки на ключевые объекты (ТОВ, Глобино-2, статья).
- Утилиты форматирования сумм и сохранения CSV/JSON.

ВАЖНО (memory feedback_balans_etalon_period_serverside):
  Даты передаём в запрос СЕРВЕРНО через ДАТАВРЕМЯ(...) встроенно в текст,
  а НЕ через q.SetParameter(datetime) — COM сдвигает на tz, окно остатков
  плывёт и НМ/КМ становятся неправильными.
"""
import sys
import json
import os
import csv

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

import win32com.client

# === Константы (зафиксированы discovery через MCP 2026-05-23) ===
EDRPOU_ORG_TOV = "40645273"           # ТОВ ІНДАСТРІАЛБУД (UUID 80d3000c…)
CODE_PODR_GLOBINO2 = "00-000298"      # Глобино-2 (UUID 9d2c9166-…)
CODE_STATYA_ZAD_KLIENTOV = "00000000002"

# Период расследования + эталон (НМ_дек = КМ_ноя)
DATES = {
    "НачНоя": "ДАТАВРЕМЯ(2025,11,1,0,0,0)",
    "КонНоя": "ДАТАВРЕМЯ(2025,11,30,23,59,59)",
    "НачДек": "ДАТАВРЕМЯ(2025,12,1,0,0,0)",
    "КонДек": "ДАТАВРЕМЯ(2025,12,31,23,59,59)",
    "НачЯнв": "ДАТАВРЕМЯ(2026,1,1,0,0,0)",
}

CONN_STRING = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

ARTIFACTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
os.makedirs(ARTIFACTS_DIR, exist_ok=True)


def connect_erp():
    """Подключение к ERP (BaseERP) через V83.COMConnector."""
    v8 = win32com.client.Dispatch("V83.COMConnector")
    erp = v8.Connect(CONN_STRING)
    return erp


def get_refs(erp):
    """Возвращает dict со ссылками на ключевые объекты.

    Поиск через НайтиПоКоду/НайтиПоРеквизиту — без UUID, чтобы избежать
    проблем парсинга UUID без дефисов в Python COM.
    """
    org = erp.Справочники.Организации.НайтиПоРеквизиту(
        "КодПоЕДРПОУ", EDRPOU_ORG_TOV
    )
    if not _has_ref(erp, org):
        raise RuntimeError(f"Организация ЕДРПОУ={EDRPOU_ORG_TOV} не найдена")

    podr = erp.Справочники.СтруктураПредприятия.НайтиПоКоду(CODE_PODR_GLOBINO2)
    if not _has_ref(erp, podr):
        raise RuntimeError(f"Подразделение Код={CODE_PODR_GLOBINO2} не найдено")

    statya = erp.ПланыВидовХарактеристик.СтатьиАктивовПассивов.НайтиПоКоду(
        CODE_STATYA_ZAD_KLIENTOV
    )
    if not _has_ref(erp, statya):
        raise RuntimeError(f"Статья Код={CODE_STATYA_ZAD_KLIENTOV} не найдена")

    istochnik = erp.Перечисления.ИсточникиУправленческогоБаланса.РасчетыСКлиентамиПоСрокам
    nashe_predpr = erp.Справочники.Партнеры.НашеПредприятие

    return {
        "Орг": org,
        "Подр": podr,
        "Статья": statya,
        "Источник": istochnik,
        "НашеПредприятие": nashe_predpr,
    }


def _has_ref(erp, ref):
    """ИспользуетсяZначениеЗаполнено по семантике 1С."""
    try:
        return bool(erp.ЗначениеЗаполнено(ref))
    except Exception:
        return ref is not None


def money(x):
    """Форматирует число как 12 345 678,90 (рус. разделители)."""
    if x is None:
        return ""
    try:
        s = f"{float(x):,.2f}"
        return s.replace(",", " ").replace(".", ",")
    except (TypeError, ValueError):
        return str(x)


def save_json(name, data):
    """Сохраняет JSON в _artifacts/<name>.json."""
    path = os.path.join(ARTIFACTS_DIR, f"{name}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    return path


def load_json(name):
    """Читает JSON из _artifacts/<name>.json."""
    path = os.path.join(ARTIFACTS_DIR, f"{name}.json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_csv(name, rows, headers):
    """Сохраняет CSV (utf-8-sig для Excel)."""
    path = os.path.join(ARTIFACTS_DIR, f"{name}.csv")
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, delimiter=";")
        w.writeheader()
        for r in rows:
            w.writerow(r)
    return path


def table_to_list(com_table):
    """ТаблицаЗначений → список dict (по именам колонок)."""
    rows = []
    cols = [c.Имя for c in com_table.Колонки]
    for i in range(com_table.Количество()):
        row = com_table.Получить(i)
        rows.append({c: row[c] for c in cols})
    return rows


def get_uuid(erp, ref):
    """Извлечь UUID ссылки через серверный erp.String(UUID).

    Прямой str() на COM-объект УникальныйИдентификатор возвращает
    '<COMObject <unknown>>' — нужно делегировать сериализацию в 1С.
    Шаблон из memory broken_refs_detection / kazna_erp_uuid_mapping.
    """
    try:
        if ref is None:
            return ""
        if not erp.ЗначениеЗаполнено(ref):
            return ""
        return str(erp.String(ref.УникальныйИдентификатор()))
    except Exception:
        return ""


def get_type_name(erp, ref):
    """Имя метаданных типа ref (например 'РеализацияТоваровУслуг').

    Через ref.Метаданные().Имя — корректно работает для непустых ссылок.
    """
    try:
        if ref is None or not erp.ЗначениеЗаполнено(ref):
            return ""
        meta = ref.Метаданные()
        return str(meta.Имя)
    except Exception:
        return ""


# ============================================================
# Утилиты для будущих скриптов (добавлены 2026-05-23 при дедупе)
# ============================================================

def find_doc_by_number(erp, doc_type, number):
    """Найти документ по Номеру через запрос (без UUID).

    Пример: find_doc_by_number(erp, "ВводОстатков", "0Ц-00000083")

    Возвращает Ссылка или None.
    """
    q = erp.NewObject("Запрос")
    q.Текст = f'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.{doc_type} ГДЕ Номер = "{number}"'
    try:
        sel = q.Выполнить().Выбрать()
        if sel.Следующий():
            return sel.Ссылка
    except Exception:
        pass
    return None


def repost_doc(erp, ref, verbose=False):
    """Перепровести документ через COM.

    Возвращает (ok: bool, msg: str).

    ВНИМАНИЕ (memory feedback_com_repost_skips_registrator_raschetov):
    НЕ перепроводить через COM документы которые пишут в РСКПС/РСППС
    в новой архитектуре взаиморасчётов — РегистраторРасчётов не создастся.
    Для таких документов — только UI 1С Enterprise.

    Безопасные категории для COM:
    - А_ФинРез_Баланс, А_ФинРез_PL (свод-документы)
    - ВводОстатков (если без изменения ТЧ-связанных регистров)
    - Корректировка регистров
    """
    obj = ref.ПолучитьОбъект()
    if obj is None:
        return False, "obj=None (битая ссылка)"
    try:
        obj.Записать(erp.РежимЗаписиДокумента.Проведение)
        if verbose:
            print(f"  [OK] {erp.String(ref)}")
        return True, "OK"
    except Exception as e:
        info = getattr(e, "excepinfo", None)
        msg = info[2] if info else str(e)
        if verbose:
            print(f"  [FAIL] {erp.String(ref)}: {msg}")
        return False, str(msg)


def find_finrez_balans(erp, org_ref, month_date):
    """Найти проведённый А_ФинРез_Баланс по (Орг, Месяц).

    month_date — datetime начала месяца (например datetime(2025,12,1)).
    Возвращает Ссылка или None.
    """
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Орг", org_ref)
    q.УстановитьПараметр("Мес", month_date)
    q.Текст = """
    ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка КАК Ссылка
    ИЗ Документ.А_ФинРез_Баланс КАК Д
    ГДЕ Д.Организация = &Орг И Д.Месяц = &Мес
        И Д.Проведен И НЕ Д.ПометкаУдаления
    """
    try:
        sel = q.Выполнить().Выбрать()
        if sel.Следующий():
            return sel.Ссылка
    except Exception:
        pass
    return None


def dump_tc(erp, obj, tc_name):
    """Дамп ТЧ объекта в список dict (для backup в JSON).

    Возвращает (cols, rows) где rows = [{col: value|{name,uuid,type}}].
    Используй для backup ПЕРЕД изменением ТЧ.
    """
    tc = getattr(obj, tc_name)
    md_tc = obj.Метаданные().ТабличныеЧасти.Найти(tc_name)
    if md_tc is None:
        return [], []
    cols = [r.Имя for r in md_tc.Реквизиты]
    rows = []
    for i in range(tc.Количество()):
        row = tc.Получить(i)
        d = {}
        for col in cols:
            try:
                v = getattr(row, col)
                if v is None:
                    d[col] = None
                elif isinstance(v, (str, int, float, bool)):
                    d[col] = v
                else:
                    try:
                        name = str(erp.String(v)) if erp.ЗначениеЗаполнено(v) else ""
                        uuid = ""
                        try:
                            if erp.ЗначениеЗаполнено(v):
                                uuid = str(erp.String(v.УникальныйИдентификатор()))
                        except Exception:
                            pass
                        d[col] = {"name": name, "uuid": uuid}
                    except Exception:
                        d[col] = "<obj>"
            except Exception:
                pass
        rows.append(d)
    return cols, rows


def query_to_rows(com_result):
    """ТаблицаЗначений (результат Выгрузить()) → list[dict] по именам колонок.

    Альтернатива table_to_list (та же логика, синоним для удобства).
    """
    return table_to_list(com_result)
