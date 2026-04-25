"""
Перевірка актуальності бази знань NotebookLM.
Порівнює дати модифікації джерельних файлів з датою генерації файлів знань.

Запуск: python _Rarzrabotki/Python/check_knowledge_freshness.py
"""

import os
import datetime

CONFIG_ROOT = r"C:\Configuration_downloads\BASERP25"
KNOWLEDGE_DIR = os.path.join(CONFIG_ROOT, "_Rarzrabotki", "notebook", "knowledge")

# Маппінг: файл знань -> список джерельних файлів
KNOWLEDGE_MAP = {
    "document_hidden_fields.md": {
        "type": "ЗМІННИЙ",
        "sources": [
            "Documents/ПриобретениеТоваровУслуг/Ext/ObjectModule.bsl",
            "Documents/РеализацияТоваровУслуг/Ext/ObjectModule.bsl",
            "Documents/ПриходныйКассовыйОрдер/Ext/ObjectModule.bsl",
            "Documents/РасходныйКассовыйОрдер/Ext/ObjectModule.bsl",
            "Documents/ПоступлениеБезналичныхДенежныхСредств/Ext/ObjectModule.bsl",
            "Documents/СписаниеБезналичныхДенежныхСредств/Ext/ObjectModule.bsl",
            "Documents/ЗаказКлиента/Ext/ObjectModule.bsl",
            "Documents/ЗаказПоставщику/Ext/ObjectModule.bsl",
            "Documents/ПеремещениеТоваров/Ext/ObjectModule.bsl",
            "Documents/АвансовыйОтчет/Ext/ObjectModule.bsl",
            "Documents/АктВыполненныхРабот/Ext/ObjectModule.bsl",
            "Documents/ВзаимозачетЗадолженности/Ext/ObjectModule.bsl",
            "Documents/ВозвратТоваровПоставщику/Ext/ObjectModule.bsl",
            "CommonModules/А_СобытияОбъектов/Ext/Module.bsl",
        ],
    },
    "document_posting_chain.md": {
        "type": "ЗМІННИЙ",
        "sources": [
            "CommonModules/А_СобытияОбъектов/Ext/Module.bsl",
        ],
    },
    "settlements_architecture.md": {
        "type": "ПОСТІЙНИЙ",
        "sources": [
            "CommonModules/ВзаиморасчетыСервер/Ext/Module.bsl",
            "CommonModules/ОперативныеВзаиморасчетыСервер/Ext/Module.bsl",
        ],
    },
    "exchange_erp_buhbud.md": {
        "type": "ЗМІННИЙ",
        "sources_external": [
            r"C:\Razrabotki\ExChange\Индастриал\ПравилаОбменаДаннымиЕРПБух\ExchangeRules.xml",
        ],
    },
    "exchange_erp_kazna.md": {
        "type": "ЗМІННИЙ",
        "sources_external": [
            os.path.join(CONFIG_ROOT, "_Rarzrabotki", "ExChange", "Казна", "ПравилаОбменаДаннымиКазна", "ExchangeRules.xml"),
        ],
    },
    "register_report_map.md": {
        "type": "ПОСТІЙНИЙ",
        "sources": [],
    },
    "custom_objects_guide.md": {
        "type": "ЗМІННИЙ",
        "sources": [
            "Documents/А_ОказаниеУслугМеждуПодразделениями/Ext/ObjectModule.bsl",
            "Documents/А_ОтражениеЗарплатыВУчете/Ext/ObjectModule.bsl",
            "Documents/А_БюджетМесяц/Ext/ObjectModule.bsl",
            "CommonModules/А_СобытияОбъектов/Ext/Module.bsl",
            "CommonModules/А_Привилегированный/Ext/Module.bsl",
            "CommonModules/А_Виробництво/Ext/Module.bsl",
        ],
    },
    "baserp25_knowledge.md": {
        "type": "ПОСТІЙНИЙ",
        "sources": [],
    },
}


def get_mtime(path):
    try:
        return datetime.datetime.fromtimestamp(os.path.getmtime(path))
    except OSError:
        return None


def main():
    print("=" * 60)
    print("ПЕРЕВІРКА АКТУАЛЬНОСТІ БАЗИ ЗНАНЬ NotebookLM")
    print("=" * 60)
    print()

    changed = []
    ok = []
    missing_knowledge = []

    for knowledge_file, config in KNOWLEDGE_MAP.items():
        knowledge_path = os.path.join(KNOWLEDGE_DIR, knowledge_file)
        knowledge_mtime = get_mtime(knowledge_path)

        if knowledge_mtime is None:
            missing_knowledge.append(knowledge_file)
            continue

        file_type = config["type"]
        sources = config.get("sources", [])
        sources_external = config.get("sources_external", [])

        # Check config-relative sources
        newer_sources = []
        for src in sources:
            src_path = os.path.join(CONFIG_ROOT, src)
            src_mtime = get_mtime(src_path)
            if src_mtime and src_mtime > knowledge_mtime:
                newer_sources.append((os.path.basename(src), src_mtime))

        # Check external sources
        for src_path in sources_external:
            src_mtime = get_mtime(src_path)
            if src_mtime and src_mtime > knowledge_mtime:
                newer_sources.append((os.path.basename(src_path), src_mtime))

        if newer_sources:
            changed.append((knowledge_file, file_type, knowledge_mtime, newer_sources))
        else:
            ok.append((knowledge_file, file_type, knowledge_mtime))

    # Output
    if changed:
        print("ПОТРЕБУЄ ОНОВЛЕННЯ:")
        print("-" * 60)
        for kf, ft, km, ns in changed:
            print(f"  {kf} [{ft}]")
            print(f"    Генерація: {km.strftime('%Y-%m-%d %H:%M')}")
            for src_name, src_mtime in ns:
                print(f"    ЗМІНЕНО: {src_name} ({src_mtime.strftime('%Y-%m-%d %H:%M')})")
        print()

    if missing_knowledge:
        print("НЕ СТВОРЕНО:")
        print("-" * 60)
        for kf in missing_knowledge:
            print(f"  {kf}")
        print()

    if ok:
        print("АКТУАЛЬНІ (OK):")
        print("-" * 60)
        for kf, ft, km in ok:
            print(f"  {kf} [{ft}] — генерація: {km.strftime('%Y-%m-%d %H:%M')}")
        print()

    print(f"Всього файлів знань: {len(KNOWLEDGE_MAP)}")
    print(f"  Актуальні: {len(ok)}")
    print(f"  Потребують оновлення: {len(changed)}")
    print(f"  Не створені: {len(missing_knowledge)}")


if __name__ == "__main__":
    main()
