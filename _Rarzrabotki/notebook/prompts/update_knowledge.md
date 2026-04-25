# Промт: Оновити базу знань NotebookLM

Скопіюй це повідомлення і вставь в чат Claude Code:

---

Перевір актуальність бази знань NotebookLM і онови якщо потрібно.

1. Запусти `python _Rarzrabotki/Python/check_knowledge_freshness.py` — подивись які файли змінились
2. Прочитай `_Rarzrabotki/notebook/knowledge/KNOWLEDGE_MAP.md` — там source ID і джерельні файли
3. Для кожного застарілого файлу знань:
   - Прочитай змінені джерельні файли (ObjectModule.bsl, Module.bsl, ExchangeRules.xml)
   - Проаналізуй що змінилось
   - Онови відповідний .md файл в _Rarzrabotki/notebook/knowledge/
   - Видали старе джерело: source_delete(notebook_id="3303acdb-2d7f-4879-9f13-78705ab3fb8c", source_id=<з KNOWLEDGE_MAP>)
   - Завантаж нове: notebook_add_text(notebook_id="3303acdb-2d7f-4879-9f13-78705ab3fb8c", title=<назва>, file_path=<шлях>)
   - Онови source_id та дату генерації в KNOWLEDGE_MAP.md
4. Протестуй запитом до NotebookLM що знання актуальні
