# -*- coding: utf-8 -*-
"""
Общий модуль для скриптов knowledge_НеоборотныеАктивы.

НАСЛЕДУЕТ паттерны от knowledge_Balanse_money/_common.py.
Специфика НОА:
  - Свод_ПрочиеАктивыПассивы_Прямой (один регистр РН.ПрочиеАктивыПассивы сам с собой).
  - Статьи: Основные средства, Нематериальные активы, Расходы текущего периода,
    Расходы будущих периодов, Вложения в необоротные активы, Налоги.
  - 3 ключевых документа (scope Март 2026):
      Документ.АмортизацияОС (версия 2.1)
      Документ.ПринятиеКУчетуОС (версия 2.1)
      Документ.ВнутреннееПотреблениеТоваров

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

# === Константы ===
EDRPOU_ORG_TOV = "40645273"           # ТОВ ІНДАСТРІАЛБУД

CONN_STRING = 'Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"'

ARTIFACTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_artifacts")
os.makedirs(ARTIFACTS_DIR, exist_ok=True)


def connect_erp():
    """Подключение к ERP (BaseERP) через V83.COMConnector."""
    v8 = win32com.client.Dispatch("V83.COMConnector")
    return v8.Connect(CONN_STRING)


def get_org(erp):
    """Найти Организация=ТОВ ІНДАСТРІАЛБУД."""
    org = erp.Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", EDRPOU_ORG_TOV)
    if not _has_ref(erp, org):
        raise RuntimeError(f"Организация ЕДРПОУ={EDRPOU_ORG_TOV} не найдена")
    return org


def _has_ref(erp, ref):
    try:
        return bool(erp.ЗначениеЗаполнено(ref))
    except Exception:
        return ref is not None


def find_statya_by_name(erp, name_substr):
    """Найти статью АктивПассив по подстроке наименования (точное LIKE)."""
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Имя", f"%{name_substr}%")
    q.Текст = """
    ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка
    ИЗ ПланВидовХарактеристик.СтатьиАктивовПассивов
    ГДЕ Наименование ПОДОБНО &Имя
    """
    try:
        sel = q.Выполнить().Выбрать()
        if sel.Следующий():
            return sel.Ссылка
    except Exception:
        pass
    return None


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
    q.УстановитьПараметр("Н", number)
    q.Текст = f'ВЫБРАТЬ ПЕРВЫЕ 1 Ссылка ИЗ Документ.{doc_type} ГДЕ Номер = &Н'
    try:
        sel = q.Выполнить().Выбрать()
        if sel.Следующий():
            return sel.Ссылка
    except Exception:
        pass
    return None


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


def fail(e):
    """Извлечь нормальное сообщение из COM-исключения."""
    info = getattr(e, "excepinfo", None)
    if info and len(info) >= 3:
        return info[2]
    return str(e)
