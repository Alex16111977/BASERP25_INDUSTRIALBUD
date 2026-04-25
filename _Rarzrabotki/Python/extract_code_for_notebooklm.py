"""
Extract 1C BSL code from BAS ERP 2.5 configuration and create
structured knowledge files for NotebookLM.

Output: _Rarzrabotki/notebook/knowledge/ folder with markdown files
containing structured code extracts.

Strategy:
- NotebookLM source limit ~500K chars per source
- Split into multiple files by category
- Include document/module name as context headers
- Focus on: custom (А_) objects, key documents, settlement modules, custom processors
"""

import os
import glob

CONFIG_ROOT = r"C:\Configuration_downloads\BASERP25"
OUTPUT_DIR = os.path.join(CONFIG_ROOT, "_Rarzrabotki", "notebook", "knowledge")
os.makedirs(OUTPUT_DIR, exist_ok=True)

MAX_SOURCE_SIZE = 200_000  # chars, conservative limit for NotebookLM (~500K tokens but chars != tokens)


def read_bsl(path):
    """Read a .bsl file trying multiple encodings."""
    for enc in ("utf-8-sig", "utf-8", "cp1251", "cp1252"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, UnicodeError):
            continue
    return None


def extract_document_modules(doc_names, label):
    """Extract ObjectModule.bsl and ManagerModule.bsl for given documents."""
    parts = []
    for doc_name in doc_names:
        doc_dir = os.path.join(CONFIG_ROOT, "Documents", doc_name)
        if not os.path.isdir(doc_dir):
            continue

        header = f"\n{'='*80}\n# Документ: {doc_name}\n{'='*80}\n"
        doc_content = header

        # ObjectModule
        obj_path = os.path.join(doc_dir, "Ext", "ObjectModule.bsl")
        if os.path.isfile(obj_path):
            code = read_bsl(obj_path)
            if code and len(code.strip()) > 10:
                doc_content += f"\n## МодульОбъекта (ObjectModule.bsl)\n{code}\n"

        # ManagerModule
        mgr_path = os.path.join(doc_dir, "Ext", "ManagerModule.bsl")
        if os.path.isfile(mgr_path):
            code = read_bsl(mgr_path)
            if code and len(code.strip()) > 10:
                doc_content += f"\n## МодульМенеджера (ManagerModule.bsl)\n{code}\n"

        if len(doc_content) > len(header) + 10:
            parts.append((doc_name, doc_content))

    return parts


def extract_common_modules(module_names, label):
    """Extract Module.bsl for given common modules."""
    parts = []
    for mod_name in module_names:
        mod_path = os.path.join(CONFIG_ROOT, "CommonModules", mod_name, "Ext", "Module.bsl")
        if not os.path.isfile(mod_path):
            continue
        code = read_bsl(mod_path)
        if not code or len(code.strip()) < 10:
            continue

        content = f"\n{'='*80}\n# ОбщийМодуль: {mod_name}\n{'='*80}\n"
        content += f"{code}\n"
        parts.append((mod_name, content))

    return parts


def extract_external_processors():
    """Extract ObjectModule.bsl from external processors."""
    proc_dir = os.path.join(CONFIG_ROOT, "_Rarzrabotki", "Обработки")
    parts = []
    if not os.path.isdir(proc_dir):
        return parts

    for entry in sorted(os.listdir(proc_dir)):
        entry_path = os.path.join(proc_dir, entry)
        if not os.path.isdir(entry_path):
            continue
        obj_path = os.path.join(entry_path, "Ext", "ObjectModule.bsl")
        if not os.path.isfile(obj_path):
            continue
        code = read_bsl(obj_path)
        if not code or len(code.strip()) < 50:
            continue

        content = f"\n{'='*80}\n# ВнешняяОбработка: {entry}\n{'='*80}\n"
        content += f"{code}\n"
        parts.append((entry, content))

    return parts


def split_large_content(name, content, max_size):
    """Split a single large module content into smaller parts by procedure boundaries."""
    if len(content) <= max_size:
        return [(name, content)]

    # Try to split by procedure/function boundaries
    lines = content.split('\n')
    parts = []
    current_lines = []
    current_size = 0
    part_num = 1
    header_size = 200  # approximate header size

    for line in lines:
        line_size = len(line) + 1
        if current_size + line_size > max_size - header_size and current_lines:
            # Try to find a good split point (Процедура/Функция boundary)
            split_idx = len(current_lines)
            for i in range(len(current_lines) - 1, max(0, len(current_lines) - 50), -1):
                l = current_lines[i].strip()
                if l.startswith('Процедура ') or l.startswith('Функция ') or l.startswith('// ==='):
                    split_idx = i
                    break

            chunk_lines = current_lines[:split_idx] if split_idx < len(current_lines) else current_lines
            remainder = current_lines[split_idx:] if split_idx < len(current_lines) else []

            chunk_text = '\n'.join(chunk_lines)
            part_name = f"{name} (часть {part_num})"
            part_content = f"\n{'='*80}\n# {part_name}\n{'='*80}\n{chunk_text}\n"
            parts.append((part_name, part_content))

            current_lines = remainder
            current_size = sum(len(l) + 1 for l in remainder)
            part_num += 1
        else:
            current_lines.append(line)
            current_size += line_size

    if current_lines:
        chunk_text = '\n'.join(current_lines)
        part_name = f"{name} (часть {part_num})" if part_num > 1 else name
        part_content = f"\n{'='*80}\n# {part_name}\n{'='*80}\n{chunk_text}\n"
        parts.append((part_name, part_content))

    return parts


def write_chunked_files(parts, base_filename, title_prefix):
    """Write parts to files, splitting if over MAX_SOURCE_SIZE."""
    # First, split any oversized individual parts
    expanded_parts = []
    for name, content in parts:
        if len(content) > MAX_SOURCE_SIZE:
            sub_parts = split_large_content(name, content, MAX_SOURCE_SIZE - 1000)
            expanded_parts.extend(sub_parts)
        else:
            expanded_parts.append((name, content))

    files_written = []
    current_content = f"# {title_prefix}\n\n"
    current_chunk = 1
    current_items = []

    for name, content in expanded_parts:
        if len(current_content) + len(content) > MAX_SOURCE_SIZE and len(current_items) > 0:
            fname = f"{base_filename}_part{current_chunk}.md"
            fpath = os.path.join(OUTPUT_DIR, fname)
            header_info = f"Содержит: {', '.join(current_items)}\n\n"
            final = current_content[:len(f"# {title_prefix}\n\n")] + header_info + current_content[len(f"# {title_prefix}\n\n"):]
            with open(fpath, "w", encoding="utf-8") as f:
                f.write(final)
            files_written.append((fpath, len(final), current_items.copy()))
            print(f"  Written: {fname} ({len(final):,} chars, {len(current_items)} items)")

            current_content = f"# {title_prefix} (часть {current_chunk + 1})\n\n"
            current_chunk += 1
            current_items = []

        current_content += content
        current_items.append(name)

    if current_items:
        fname = f"{base_filename}_part{current_chunk}.md" if current_chunk > 1 else f"{base_filename}.md"
        fpath = os.path.join(OUTPUT_DIR, fname)
        header_info = f"Содержит: {', '.join(current_items)}\n\n"
        final = current_content[:len(f"# {title_prefix}")+5] + header_info + current_content[len(f"# {title_prefix}")+5:]
        with open(fpath, "w", encoding="utf-8") as f:
            f.write(final)
        files_written.append((fpath, len(final), current_items.copy()))
        print(f"  Written: {fname} ({len(final):,} chars, {len(current_items)} items)")

    return files_written


def main():
    all_files = []

    # 1. Key documents - ObjectModule (posting logic)
    print("\n=== 1. Key Document Modules ===")
    # Only documents actually used in the database (checked via MCP query 2026-04-08)
    key_docs = [
        "ПриобретениеТоваровУслуг",     # 2965 docs
        "ЗаказПоставщику",               # 8438 docs
        "РасходныйКассовыйОрдер",        # 1378 docs
        "ПеремещениеТоваров",             # 1059 docs
        "ЗаказКлиента",                   # 685 docs
        "ПриходныйКассовыйОрдер",        # 676 docs
        "РеализацияТоваровУслуг",         # 221 docs
        "АвансовыйОтчет",                 # 194 docs
    ]
    parts = extract_document_modules(key_docs, "Key Documents")
    files = write_chunked_files(parts, "code_key_documents", "Код ключевых документов BAS ERP 2.5 (ObjectModule + ManagerModule)")
    all_files.extend(files)

    # 2. Custom documents (А_)
    print("\n=== 2. Custom Document Modules (А_) ===")
    custom_docs = [d for d in os.listdir(os.path.join(CONFIG_ROOT, "Documents"))
                   if d.startswith("А_") and not d.endswith(".xml")]
    parts = extract_document_modules(custom_docs, "Custom Documents")
    files = write_chunked_files(parts, "code_custom_documents", "Код кастомных документов (А_) BAS ERP 2.5")
    all_files.extend(files)

    # 3. Key common modules (settlement, posting, custom)
    print("\n=== 3. Key Common Modules ===")
    key_modules = [
        "А_СобытияОбъектов",
        "А_Виробництво",
        "А_Привилегированный",
        "ВзаиморасчетыСервер",
        "ОперативныеВзаиморасчетыСервер",
        "ПроведениеСервер",
        "ПроведениеСерверУТ",
        "ОбщегоНазначения",
        "ДенежныеСредстваСервер",
        "РасчетыСПоставщикамиСервер",
        "РасчетыСКлиентамиСервер",
        "ТоварыНаСкладахСервер",
        "ВзаиморасчетыФормы",
        "ВзаиморасчетыКлиентСервер",
        "ФинансовыйУчетСервер",
        "ЗарплатаКадрыОбщиеНаборыДанных",
    ]
    parts = extract_common_modules(key_modules, "Key Common Modules")
    files = write_chunked_files(parts, "code_common_modules", "Код ключевых общих модулей BAS ERP 2.5")
    all_files.extend(files)

    # 4. External processors (custom)
    print("\n=== 4. External Processors ===")
    parts = extract_external_processors()
    files = write_chunked_files(parts, "code_external_processors", "Код внешних обработок (_Rarzrabotki/Обработки/)")
    all_files.extend(files)

    # Summary
    print(f"\n{'='*60}")
    print(f"TOTAL FILES: {len(all_files)}")
    total_chars = sum(sz for _, sz, _ in all_files)
    total_modules = sum(len(items) for _, _, items in all_files)
    print(f"TOTAL SIZE: {total_chars:,} chars")
    print(f"TOTAL MODULES: {total_modules}")
    print(f"\nFiles created in: {OUTPUT_DIR}")
    for fpath, sz, items in all_files:
        print(f"  {os.path.basename(fpath)}: {sz:,} chars ({len(items)} modules)")


if __name__ == "__main__":
    main()
