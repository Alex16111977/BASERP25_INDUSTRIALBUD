# -*- coding: utf-8 -*-
"""
Общий модуль для скриптов knowledge_Balanse_money.

НАСЛЕДУЕТ паттерны от knowledge_Balanse_klient/_common.py
+ добавляет специфику для денег (4 регистра РНДС, эталон от Отчёт.УправленческийБаланс).

ВАЖНО (memory feedback_balans_etalon_period_serverside):
  Даты передаём в запрос СЕРВЕРНО через ДАТАВРЕМЯ(...) в тексте,
  а НЕ через q.SetParameter(datetime).
"""
import sys
import json
import os
import csv

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

import win32com.client

# === Константы (зафиксированы 2026-05-23) ===
EDRPOU_ORG_TOV = "40645273"           # ТОВ ІНДАСТРІАЛБУД

# Статьи ПВХ.СтатьиАктивовПассивов (по коду)
CODE_STATYA_DS_BEZNAL = "00000000003"
CODE_STATYA_DS_NALICH = "00000000004"
CODE_STATYA_DS_PODOTCH = "00000000005"
CODE_STATYA_DS_VPUTI = "00000000028"

CONN_STRING = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

ARTIFACTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
os.makedirs(ARTIFACTS_DIR, exist_ok=True)


def connect_erp():
    """Подключение к ERP (BaseERP) через V83.COMConnector."""
    v8 = win32com.client.Dispatch("V83.COMConnector")
    return v8.Connect(CONN_STRING)


def get_refs(erp):
    """Возвращает dict со ссылками на ключевые объекты для денег.

    Перечисления.ИсточникиУправленческогоБаланса:
      ДенежныеСредстваБезналичные, ДенежныеСредстваНаличные,
      ДенежныеСредстваВПути, ДенежныеСредстваУПодотчетныхЛиц
    """
    org = erp.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", EDRPOU_ORG_TOV)
    if not _has_ref(erp, org):
        raise RuntimeError(f"Организация ЕДРПОУ={EDRPOU_ORG_TOV} не найдена")

    return {
        "Орг": org,
        "Стат_Безнал": _find_statya(erp, CODE_STATYA_DS_BEZNAL),
        "Стат_Налич":  _find_statya(erp, CODE_STATYA_DS_NALICH),
        "Стат_Подотч": _find_statya(erp, CODE_STATYA_DS_PODOTCH),
        "Стат_ВПути":  _find_statya(erp, CODE_STATYA_DS_VPUTI),
        "Ист_Безнал":  erp.Перечисления.ИсточникиУправленческогоБаланса.ДенежныеСредстваБезналичные,
        "Ист_Налич":   erp.Перечисления.ИсточникиУправленческогоБаланса.ДенежныеСредстваНаличные,
        "Ист_Подотч":  erp.Перечисления.ИсточникиУправленческогоБаланса.ДенежныеСредстваУПодотчетныхЛиц,
        "Ист_ВПути":   erp.Перечисления.ИсточникиУправленческогоБаланса.ДенежныеСредстваВПути,
    }


def _find_statya(erp, code):
    s = erp.ПланыВидовХарактеристик.СтатьиАктивовПассивов.НайтиПоКоду(code)
    return s if _has_ref(erp, s) else None


def _has_ref(erp, ref):
    try:
        return bool(erp.ЗначениеЗаполнено(ref))
    except Exception:
        return ref is not None


def find_podr_by_code(erp, code):
    """Найти подразделение по коду (например '0Ц-000007' = ЦО)."""
    return erp.Справочники.СтруктураПредприятия.НайтиПоКоду(code)


def money(x):
    """Форматирует число как 12 345 678,90."""
    if x is None:
        return ""
    try:
        s = f"{float(x):,.2f}"
        return s.replace(",", " ").replace(".", ",")
    except (TypeError, ValueError):
        return str(x)


def save_json(name, data):
    path = os.path.join(ARTIFACTS_DIR, f"{name}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    return path


def load_json(name):
    path = os.path.join(ARTIFACTS_DIR, f"{name}.json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_csv(name, rows, headers):
    path = os.path.join(ARTIFACTS_DIR, f"{name}.csv")
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, delimiter=";")
        w.writeheader()
        for r in rows:
            w.writerow(r)
    return path


def get_uuid(erp, ref):
    """Извлечь UUID ссылки через серверный erp.String(UUID)."""
    try:
        if ref is None or not erp.ЗначениеЗаполнено(ref):
            return ""
        return str(erp.String(ref.УникальныйИдентификатор()))
    except Exception:
        return ""


def get_type_name(erp, ref):
    """Имя метаданных типа ref."""
    try:
        if ref is None or not erp.ЗначениеЗаполнено(ref):
            return ""
        return str(ref.Метаданные().Имя)
    except Exception:
        return ""


def find_doc_by_number(erp, doc_type, number):
    """Найти документ по Номеру (для перепроведения / drill-down)."""
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
    """Перепровести документ. (ok, msg)"""
    obj = ref.ПолучитьОбъект()
    if obj is None:
        return False, "obj=None"
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
    """Найти проведённый А_ФинРез_Баланс по (Орг, Месяц)."""
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


def table_to_list(com_table):
    """ТаблицаЗначений → список dict."""
    rows = []
    cols = [c.Имя for c in com_table.Колонки]
    for i in range(com_table.Количество()):
        row = com_table.Получить(i)
        rows.append({c: row[c] for c in cols})
    return rows
