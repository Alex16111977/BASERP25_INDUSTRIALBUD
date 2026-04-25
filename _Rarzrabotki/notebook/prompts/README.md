# Промти для роботи з NotebookLM Knowledge Base

Notebook ID: 3303acdb-2d7f-4879-9f13-78705ab3fb8c
Notebook: BAS ERP 2.5 INDUSTRIALBUD

## Як використовувати

Відкрий потрібний файл, скопіюй текст після "---" і вставь в чат Claude Code.

## Список промтів

| Файл | Коли використовувати |
|------|---------------------|
| update_knowledge.md | "Обнови знання" — перевірити актуальність і оновити застарілі файли |
| consult_before_coding.md | Перед складною задачею — запитати NotebookLM про ризики |
| diagnose_sync_problem.md | Коли не сходяться залишки між базами |
| add_new_knowledge.md | Коли знайшов щось нове що треба зберегти |
| create_arenda_to_okazanie.md | Створити код: А_АрендаТехники -> А_ОказаниеУслугМеждуПодразделениями |
| fix_fin_agent_doc_peredachi.md | Доробити пошук платіжки в А_ПриходДенегОтФинАгента по регістру |

## Додаткові команди

Перевірка актуальності знань:
  python _Rarzrabotki/Python/check_knowledge_freshness.py

Маніфест знань (source ID, джерельні файли):
  _Rarzrabotki/notebook/knowledge/KNOWLEDGE_MAP.md

Скіл для Claude Code:
  .claude/skills/consult-notebooklm/SKILL.md
