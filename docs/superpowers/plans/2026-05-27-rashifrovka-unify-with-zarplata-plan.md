# Унификация ТЧ.А_Расшифровка с ТЧ.Зарплата — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Привести структуру ТЧ `А_РасшифровкаВыплатыЗарплатаПоФизлицам` к единой логике с типовой ТЧ `Зарплата`: удалить `ОтработаноЧасов`, переименовать `Сумма` → `КВыплате` (с сохранением UUID), добавить 6 полей из ТЧ.Зарплата. Переписать заполнение в `Documents/РаспределениеФ2/Ext/ObjectModule.bsl` на зеркалирование 1:1. Перегенерировать форму подформы. Обновить тесты.

**Architecture:** Структурное изменение метаданных + переписывание логики заполнения. Принцип «зеркало ТЧ.Зарплата 1:1 + 2 наших специфичных поля (НаправлениеДеятельности, СтатьяДДС)». Свёртка не нужна — строки 1:1 со строками Зарплата.

**Tech Stack:** XML (метаданные документа 1С), BSL (1С 8.3.20), Python COM (V83.COMConnector + pywin32), `/form-compile` skill, db-load-xml + db-update PowerShell skills для доставки.

**Спецификация:** `docs/superpowers/specs/2026-05-27-rashifrovka-unify-with-zarplata-design.md`

---

## File Structure

| Файл | Действие | Назначение |
|---|---|---|
| `Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml` | Modify | Изменить структуру ТЧ.А_РасшифровкаВыплатыЗарплатаПоФизлицам: удалить ОтработаноЧасов, переименовать Сумма→КВыплате, добавить 6 атрибутов |
| `Documents/РаспределениеФ2/Ext/ObjectModule.bsl` | Modify | Шаг 8 — переписать на зеркалирование ТЧ.Зарплата |
| `Documents/ВедомостьНаВыплатуЗарплатыВКассу/Forms/А_ФормаСпискаРасшифровкиПоФЛ/Ext/Form.xml` | Modify | Перегенерировать через /form-compile с новым JSON DSL |
| `_Rarzrabotki/Python/test/test_create_vedomost_vkassu_iz_f2_rashifrovka.py` | Modify | Σ часов убрать, проверять Σ КВыплате |
| `_Rarzrabotki/Python/test/test_vkassu_subform_returns_changes.py` | Modify | Сумма → КВыплате; убрать ОтработаноЧасов |
| `_Rarzrabotki/Python/test/test_vkassu_subform_metadata.py` | Modify | Обновить список ожидаемых реквизитов ТЧ |
| `_Rarzrabotki/Python/scaffold_rashifrovka_struct.py` | Create (temp) | Python helper для безопасных правок XML с CRLF + кириллицей |

---

## Task 1: Pre-flight — изучить XML-структуру эталона ТЧ.Зарплата

Нам нужны точные UUIDs реквизита `Сумма` (для сохранения) и blob `<Type>` для `ДокументОснование` из ТЧ.Зарплата (15 типов composite ref — копируем буквально).

**Files:**
- Read: `Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml`

- [ ] **Step 1: Найти UUID атрибута `Сумма` в ТЧ.А_РасшифровкаВыплатыЗарплатаПоФизлицам**

Run:
```
Grep("<Name>Сумма</Name>", "Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml", output_mode="content", n=true, "-B"=2)
```

Найти блок `<Attribute uuid="..."><Properties><Name>Сумма</Name>` — внутри секции А_РасшифровкаВыплатыЗарплатаПоФизлицам (а не НДФЛ/ВзносыФОТ — в них тоже есть `<Name>Сумма</Name>`). **Зафиксировать UUID**.

Чтобы найти правильный блок:
```
Grep("name=\"А_РасшифровкаВыплатыЗарплатаПоФизлицам\"|<Name>Сумма</Name>", "Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml", output_mode="content", n=true)
```
Берём `<Name>Сумма</Name>` с номером строки **сразу после** строки `name="А_РасшифровкаВыплатыЗарплатаПоФизлицам"`.

- [ ] **Step 2: Зафиксировать структуру ТЧ.Зарплата (атрибут `ДокументОснование`)**

Run:
```
Grep("<Attribute uuid|<Name>ДокументОснование</Name>", "Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml", output_mode="content", n=true, head_limit=50)
```

Найти `<Attribute uuid="..."><Properties><Name>ДокументОснование</Name>` в **ТЧ.Зарплата** (есть также в ТЧ.НДФЛ — там 17 типов). Для ТЧ.Зарплата 15 типов. Запомнить строку начала блока.

Затем:
```
Read("Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml", offset=<строка_начала>, limit=80)
```
Видим полный блок `<Attribute>...<Type>...<v8:Type>cfg:DocumentRef.ЕдиновременноеПособиеЗаСчетФСС</v8:Type>...</Type>...</Attribute>`. **Запомнить точные строки начала и конца** (для копирования в Task 2).

- [ ] **Step 3: Прочитать существующий блок ТЧ.А_РасшифровкаВыплатыЗарплатаПоФизлицам целиком**

Run:
```
Grep("name=\"А_РасшифровкаВыплатыЗарплатаПоФизлицам\"", "Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml", output_mode="content", n=true)
```
Найти `<TabularSection uuid="..." name="А_РасшифровкаВыплатыЗарплатаПоФизлицам">`. Затем:

```
Read("Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml", offset=<строка_TabularSection>, limit=200)
```
Найти закрывающий `</TabularSection>`. Зафиксировать диапазон строк.

- [ ] **Step 4: Сгенерировать UUIDs для 6 новых атрибутов**

В Python (`mcp__python-runner__run_command`):
```python
import uuid
for name in ["ПериодВзаиморасчетов", "СтатьяФинансирования", "СтатьяРасходов",
              "ДокументОснование", "КомпенсацияЗаЗадержкуЗарплаты", "ГруппаУчетаНачислений"]:
    print(f"{name}: {uuid.uuid4()}")
```

Запомнить 6 UUIDs — используем в Task 2.

---

## Task 2: Изменить структуру ТЧ в Documents/ВКассу.xml

Через Python helper (CRLF + кириллица + длинный composite ref → ручные Edit рискованны).

**Files:**
- Create: `_Rarzrabotki/Python/scaffold_rashifrovka_struct.py` (одноразовый)
- Modify: `Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml`

- [ ] **Step 1: Создать Python helper**

Создать `_Rarzrabotki/Python/scaffold_rashifrovka_struct.py`:

```python
"""
Одноразовый: переструктурировать ТЧ.А_РасшифровкаВыплатыЗарплатаПоФизлицам в
Documents/ВКассу.xml.

Действия:
1. Найти ТЧ А_Расшифровка в XML
2. Удалить блок <Attribute><Name>ОтработаноЧасов</Name>...</Attribute>
3. Переименовать <Name>Сумма</Name> → <Name>КВыплате</Name> (UUID сохраняем)
4. Найти ТЧ.Зарплата.ДокументОснование → скопировать <Type>...</Type> блок
5. Перед закрывающим </TabularSection> ТЧ.А_Расшифровка вставить 6 новых атрибутов

UUIDs новых атрибутов — подставить из Task 1 Step 4.
"""
import io, sys, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

PATH = r"C:\Configuration_downloads\BASERP25\Documents\ВедомостьНаВыплатуЗарплатыВКассу.xml"

# UUIDs из Task 1 Step 4 — подставить
UUID_PERIOD       = "<UUID-ПериодВзаиморасчетов>"
UUID_ST_FIN       = "<UUID-СтатьяФинансирования>"
UUID_ST_RASH      = "<UUID-СтатьяРасходов>"
UUID_DOC_OSN      = "<UUID-ДокументОснование>"
UUID_KOMP         = "<UUID-КомпенсацияЗаЗадержкуЗарплаты>"
UUID_GU_NACH      = "<UUID-ГруппаУчетаНачислений>"

with open(PATH, "rb") as f:
    data = f.read().decode("utf-8")
print(f"Before: {len(data.encode('utf-8'))} bytes")

# === ПРОВЕРКА — найти строки ===
# Регулярки на крупные секции
# 1. Найти границы ТЧ А_Расшифровка
m_tch_start = re.search(r'<TabularSection uuid="[^"]+" name="А_РасшифровкаВыплатыЗарплатаПоФизлицам">', data)
if not m_tch_start:
    print("FAIL: ТЧ А_РасшифровкаВыплатыЗарплатаПоФизлицам не найдена")
    sys.exit(1)
tch_start_idx = m_tch_start.start()

# Найти закрывающий </TabularSection> для нашей ТЧ
# (искать первый </TabularSection> ПОСЛЕ tch_start_idx)
m_tch_end = re.search(r'</TabularSection>', data[tch_start_idx:])
if not m_tch_end:
    print("FAIL: </TabularSection> для А_Расшифровки не найден")
    sys.exit(1)
tch_end_idx = tch_start_idx + m_tch_end.start()
tch_block = data[tch_start_idx:tch_end_idx]
print(f"  ТЧ А_Расшифровка: символы {tch_start_idx}..{tch_end_idx}, длина {tch_end_idx - tch_start_idx}")

# === 1. Удалить блок ОтработаноЧасов ===
# Блок: <Attribute uuid="..."><Properties>...<Name>ОтработаноЧасов</Name>...</Properties>...</Attribute>
pattern_otrabotano = re.compile(
    r'\s*<Attribute uuid="[^"]+">\s*<Properties>\s*<Name>ОтработаноЧасов</Name>.*?</Attribute>',
    re.DOTALL
)
new_tch_block = pattern_otrabotano.sub('', tch_block, count=1)
if new_tch_block == tch_block:
    print("WARN: блок ОтработаноЧасов не найден или не удалён")
else:
    print("  ✓ Удалён блок ОтработаноЧасов")

# === 2. Переименовать Сумма → КВыплате (UUID сохранить) ===
# Внутри блока: <Properties>...<Name>Сумма</Name><Synonym>...«Сумма»...</Synonym>...
# Меняем только <Name>Сумма</Name> и Synonym/Comment (Synonym для UI)
# UUID реквизита остаётся
pattern_summa = re.compile(
    r'(<Attribute uuid="[^"]+">\s*<Properties>\s*)<Name>Сумма</Name>',
    re.DOTALL
)
new_tch_block_2 = pattern_summa.sub(r'\1<Name>КВыплате</Name>', new_tch_block, count=1)
if new_tch_block_2 == new_tch_block:
    print("WARN: блок Сумма не найден или не переименован")
else:
    print("  ✓ Переименовано Сумма → КВыплате (UUID сохранён)")

# Также поменять Synonym внутри этого блока
# Synonym блока «Сумма» содержит <v8:content>Сумма</v8:content> — поменяем на «К выплате»
# Делаем это аккуратно — только в первом найденном блоке после <Name>КВыплате</Name>
m_kvyplate = re.search(r'<Name>КВыплате</Name>', new_tch_block_2)
if m_kvyplate:
    pos = m_kvyplate.end()
    # Найти ближайший <Synonym>...</Synonym>
    m_syn = re.search(r'(<Synonym>.*?</Synonym>)', new_tch_block_2[pos:pos+500], re.DOTALL)
    if m_syn:
        old_syn = m_syn.group(1)
        new_syn = re.sub(r'<v8:content>Сумма</v8:content>', '<v8:content>К выплате</v8:content>', old_syn)
        new_syn = re.sub(r'<v8:content>Сума</v8:content>', '<v8:content>До виплати</v8:content>', new_syn)
        new_tch_block_2 = new_tch_block_2[:pos] + new_tch_block_2[pos:].replace(old_syn, new_syn, 1)
        print("  ✓ Synonym КВыплате обновлён")

# === 3. Найти <Type> блок ДокументОснование из ТЧ.Зарплата для копирования ===
# Глобально по data — ищем секцию ТЧ.Зарплата
m_zarplata = re.search(r'<TabularSection uuid="[^"]+" name="Зарплата">', data)
if not m_zarplata:
    print("FAIL: ТЧ.Зарплата не найдена")
    sys.exit(1)
zarp_start = m_zarplata.start()
zarp_end_match = re.search(r'</TabularSection>', data[zarp_start:])
zarp_end = zarp_start + zarp_end_match.start()
zarp_block = data[zarp_start:zarp_end]

# Найти атрибут ДокументОснование внутри ТЧ.Зарплата
m_doc_osn = re.search(
    r'<Attribute uuid="[^"]+">\s*<Properties>\s*<Name>ДокументОснование</Name>.*?</Attribute>',
    zarp_block, re.DOTALL
)
if not m_doc_osn:
    print("FAIL: ДокументОснование в ТЧ.Зарплата не найден")
    sys.exit(1)
doc_osn_template = m_doc_osn.group(0)
# Заменить UUID на наш
doc_osn_new = re.sub(r'<Attribute uuid="[^"]+">', f'<Attribute uuid="{UUID_DOC_OSN}">', doc_osn_template, count=1)
print(f"  ✓ Шаблон ДокументОснование найден ({len(doc_osn_template)} символов)")

# === 4. Сформировать 6 новых атрибутов ===
def make_attr(uuid_val, name, ru_title, uk_title, type_xml):
    """Простой атрибут с одним типом."""
    return f'''\t\t\t<Attribute uuid="{uuid_val}">
\t\t\t\t<Properties>
\t\t\t\t\t<Name>{name}</Name>
\t\t\t\t\t<Synonym>
\t\t\t\t\t\t<v8:item>
\t\t\t\t\t\t\t<v8:lang>ru</v8:lang>
\t\t\t\t\t\t\t<v8:content>{ru_title}</v8:content>
\t\t\t\t\t\t</v8:item>
\t\t\t\t\t\t<v8:item>
\t\t\t\t\t\t\t<v8:lang>uk</v8:lang>
\t\t\t\t\t\t\t<v8:content>{uk_title}</v8:content>
\t\t\t\t\t\t</v8:item>
\t\t\t\t\t</Synonym>
\t\t\t\t\t<Comment/>
\t\t\t\t\t<Type>
{type_xml}
\t\t\t\t\t</Type>
\t\t\t\t\t<ToolTip/>
\t\t\t\t\t<FillFromFillingValue>false</FillFromFillingValue>
\t\t\t\t\t<FillChecking>DontCheck</FillChecking>
\t\t\t\t\t<FullTextSearch>Use</FullTextSearch>
\t\t\t\t\t<DataHistory>Use</DataHistory>
\t\t\t\t</Properties>
\t\t\t</Attribute>'''

type_period = '''\t\t\t\t\t\t<v8:Type>xs:dateTime</v8:Type>
\t\t\t\t\t\t<v8:DateQualifiers>
\t\t\t\t\t\t\t<v8:DateFractions>DateTime</v8:DateFractions>
\t\t\t\t\t\t</v8:DateQualifiers>'''

type_ref = lambda full: f'\t\t\t\t\t\t<v8:Type>{full}</v8:Type>'

type_decimal = '''\t\t\t\t\t\t<v8:Type>xs:decimal</v8:Type>
\t\t\t\t\t\t<v8:NumberQualifiers>
\t\t\t\t\t\t\t<v8:Digits>15</v8:Digits>
\t\t\t\t\t\t\t<v8:FractionDigits>2</v8:FractionDigits>
\t\t\t\t\t\t\t<v8:AllowedSign>Any</v8:AllowedSign>
\t\t\t\t\t\t</v8:NumberQualifiers>'''

new_attrs = [
    make_attr(UUID_PERIOD, "ПериодВзаиморасчетов", "Период взаиморасчетов", "Період взаєморозрахунків", type_period),
    make_attr(UUID_ST_FIN, "СтатьяФинансирования", "Статья финансирования", "Стаття фінансування",
              type_ref("cfg:CatalogRef.СтатьиФинансированияЗарплата")),
    make_attr(UUID_ST_RASH, "СтатьяРасходов", "Статья расходов", "Стаття витрат",
              type_ref("cfg:CatalogRef.СтатьиРасходовЗарплата")),
    doc_osn_new,  # уже сформирован выше (15 типов)
    make_attr(UUID_KOMP, "КомпенсацияЗаЗадержкуЗарплаты", "Компенсация за задержку зарплаты",
              "Компенсація за затримку зарплати", type_decimal),
    make_attr(UUID_GU_NACH, "ГруппаУчетаНачислений", "Группа учета начислений", "Група обліку нарахувань",
              type_ref("cfg:CatalogRef.ГруппыУчетаНачисленийИУдержаний")),
]

new_attrs_block = "\n" + "\n".join(new_attrs) + "\n\t\t"

# === 5. Вставить новые атрибуты перед закрывающим </TabularSection> ===
# Найти позицию: ищем последний </Attribute> в new_tch_block_2, после него вставляем
m_last_attr = None
for m in re.finditer(r'</Attribute>', new_tch_block_2):
    m_last_attr = m
if not m_last_attr:
    print("FAIL: ни одного </Attribute> в ТЧ.А_Расшифровка")
    sys.exit(1)
insert_pos = m_last_attr.end()
final_tch_block = new_tch_block_2[:insert_pos] + new_attrs_block + new_tch_block_2[insert_pos:]
print(f"  ✓ Вставлено 6 новых атрибутов на позицию {insert_pos}")

# === 6. Подставить обратно в data ===
new_data = data[:tch_start_idx] + final_tch_block + data[tch_end_idx:]
# Сохранить с CRLF
# Если data уже было LF-only — оставить LF; если CRLF — сохранить CRLF
# Проверяем самый частый стиль
crlf_count = data.count("\r\n")
lf_only = data.count("\n") - crlf_count
print(f"  Detected: \\r\\n={crlf_count}, lone \\n={lf_only}")
if crlf_count > lf_only:
    # CRLF style — наши вставки используют \n → конвертируем
    new_data = re.sub(r'(?<!\r)\n', '\r\n', new_data)

with open(PATH, "wb") as f:
    f.write(new_data.encode("utf-8"))
print(f"After: {len(new_data.encode('utf-8'))} bytes")
print("DONE")
```

ВАЖНО: перед запуском подставить 6 UUIDs из Task 1 Step 4 в верх скрипта.

- [ ] **Step 2: Подставить UUIDs из Task 1 Step 4 в скрипт**

Через Edit заменить 6 placeholder-строк `"<UUID-...>"` на реальные UUIDs.

- [ ] **Step 3: Запустить scaffold**

Run:
```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\scaffold_rashifrovka_struct.py"
```

Expected:
```
Before: <N> bytes
  ТЧ А_Расшифровка: символы ...
  ✓ Удалён блок ОтработаноЧасов
  ✓ Переименовано Сумма → КВыплате (UUID сохранён)
  ✓ Synonym КВыплате обновлён
  ✓ Шаблон ДокументОснование найден (... символов)
  ✓ Вставлено 6 новых атрибутов на позицию ...
After: <N + ~5000> bytes
DONE
```

- [ ] **Step 4: Проверить результат через grep**

Run:
```
Grep("<Name>(ОтработаноЧасов|Сумма|КВыплате|ПериодВзаиморасчетов|КомпенсацияЗаЗадержкуЗарплаты)</Name>", "Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml", output_mode="content", n=true, glob="*.xml")
```

Expected:
- `<Name>ОтработаноЧасов</Name>` — **отсутствует** (в нашей ТЧ — других нет)
- `<Name>Сумма</Name>` — должны быть **только в других ТЧ** (НДФЛ, ВзносыФОТ), не в А_Расшифровке
- `<Name>КВыплате</Name>` — **2 строки** (одна в ТЧ.Зарплата, одна в А_Расшифровке)
- `<Name>ПериодВзаиморасчетов</Name>` — **должно быть N+1** строк (была в Зарплата, НДФЛ, ВзносыФОТ; теперь ещё в А_Расшифровке)
- `<Name>КомпенсацияЗаЗадержкуЗарплаты</Name>` — **2 строки** (Зарплата и А_Расшифровка)

- [ ] **Step 5: Удалить scaffold script**

Run:
```bash
rm "_Rarzrabotki/Python/scaffold_rashifrovka_struct.py"
```

- [ ] **Step 6: Commit структуры**

```bash
git add "Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml"
git commit -m "$(cat <<'EOF'
feat(rashifrovka-unify): структура ТЧ.А_Расшифровка приведена к ТЧ.Зарплата

Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml — изменена ТЧ
А_РасшифровкаВыплатыЗарплатаПоФизлицам:
- удалён реквизит ОтработаноЧасов
- переименован Сумма → КВыплате (UUID сохранён → данные мигрируют)
- добавлены 6 новых реквизитов из ТЧ.Зарплата:
  - ПериодВзаиморасчетов (Дата)
  - СтатьяФинансирования (СтатьиФинансированияЗарплата)
  - СтатьяРасходов (СтатьиРасходовЗарплата)
  - ДокументОснование (composite 15 типов как в Зарплата)
  - КомпенсацияЗаЗадержкуЗарплаты (Число)
  - ГруппаУчетаНачислений (ГруппыУчетаНачисленийИУдержаний)

Итого 13 полей: 11 из них зеркало ТЧ.Зарплата + 2 наших
(НаправлениеДеятельности, СтатьяДвиженияДенежныхСредств).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 3: Переписать Шаг 8 в РаспределениеФ2/Ext/ObjectModule.bsl

**Files:**
- Modify: `Documents/РаспределениеФ2/Ext/ObjectModule.bsl`

- [ ] **Step 1: Прочитать текущий Шаг 8**

Run:
```
Grep("// === 8\\. ТЧ А_Расшифровка|Свернуть\\(", "Documents/РаспределениеФ2/Ext/ObjectModule.bsl", output_mode="content", n=true, "-A"=2)
```

Запомнить точные строки начала и конца Шага 8 + блок Свернуть.

- [ ] **Step 2: Прочитать полный блок Шага 8**

Run:
```
Read("Documents/РаспределениеФ2/Ext/ObjectModule.bsl", offset=<строка_Шаг_8>, limit=80)
```

Зафиксировать что заменяем — от заголовка `// === 8. ...` до строки перед `// === 9. Сумма по документу ===` (если он есть в коде).

- [ ] **Step 3: Заменить Шаг 8 через Edit**

Edit (точные тексты — `old_string` = текущий Шаг 8 включая Свернуть; `new_string` = новый блок):

Старый (общий вид — точные строки сверяются с файлом):
```bsl
	// === 8. ТЧ А_РасшифровкаВыплатыЗарплатаПоФизлицам ← ТЧ Распределение Ф2 ===
	Для Каждого СтрР Из Распределение Цикл
		... [старая логика по СтрР]
		НовСтр.ОтработаноЧасов               = СтрР.ОтработаноЧасов;
		НовСтр.Сумма                         = СтрР.СуммаНачисления;
		...
	КонецЦикла;

	// Свёртка дубликатов ключа
	ВыкВКассу.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Свернуть(
		"ФизическоеЛицо, Сотрудник, Подразделение, НаправлениеДеятельности, СтатьяДвиженияДенежныхСредств",
		"ОтработаноЧасов, Сумма");
```

Новый:
```bsl
	// === 8. ТЧ А_Расшифровка — зеркало ТЧ.Зарплата 1:1 + СтатьяДДС + Направление ===
	// Поля 1-11 копируются из строки Зарплата (ИдентификаторСтроки, Сотрудник, ФизическоеЛицо,
	// Подразделение, ПериодВзаиморасчетов, СтатьяФинансирования, СтатьяРасходов,
	// ДокументОснование, КВыплате, КомпенсацияЗаЗадержкуЗарплаты, ГруппаУчетаНачислений).
	// Поля 12-13 (НаправлениеДеятельности, СтатьяДвиженияДенежныхСредств) — наши специфичные.
	ПоляЗеркала = "ИдентификаторСтроки, Сотрудник, ФизическоеЛицо, Подразделение,"
		+ " ПериодВзаиморасчетов, СтатьяФинансирования, СтатьяРасходов,"
		+ " ДокументОснование, КВыплате, КомпенсацияЗаЗадержкуЗарплаты, ГруппаУчетаНачислений";

	Для Каждого СтрЗ Из ВыкВКассу.Зарплата Цикл
		НовСтр = ВыкВКассу.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Добавить();
		ЗаполнитьЗначенияСвойств(НовСтр, СтрЗ, ПоляЗеркала);

		// НаправлениеДеятельности из шапки Ф2.Направление
		НовСтр.НаправлениеДеятельности = Направление;

		// СтатьяДДС: 1) из словаря СтатьиДДСПоКлючу (новая кнопка ОтражениеЗПпоКазне),
		//             2) иначе fallback на Сотрудник.СтатьяЗарплата → Родитель.СтатьяЗарплата.
		Если СтатьиДДСПоКлючу <> Неопределено Тогда
			Ключ_СтД = ?(ЗначениеЗаполнено(СтрЗ.Сотрудник),
						Строка(СтрЗ.Сотрудник.УникальныйИдентификатор()), "")
				+ "|"
				+ ?(ЗначениеЗаполнено(СтрЗ.Подразделение),
						Строка(СтрЗ.Подразделение.УникальныйИдентификатор()), "");
			СтатьяДДС = СтатьиДДСПоКлючу.Получить(Ключ_СтД);
			Если СтатьяДДС = Неопределено ИЛИ НЕ ЗначениеЗаполнено(СтатьяДДС) Тогда
				СтатьяДДС = Справочники.СтатьиДвиженияДенежныхСредств.ПустаяСсылка();
				Сообщить("WARN: СтатьяДДС не найдена в ОтражениеЗПпоКазне для "
					+ Строка(СтрЗ.Сотрудник) + " / " + Строка(СтрЗ.Подразделение));
			КонецЕсли;
		Иначе
			СтатьяДДС = СтрЗ.Сотрудник.СтатьяЗарплата;
			Если НЕ ЗначениеЗаполнено(СтатьяДДС) Тогда
				Попытка
					СтатьяДДС = СтрЗ.Сотрудник.Родитель.СтатьяЗарплата;
				Исключение
					СтатьяДДС = Справочники.СтатьиДвиженияДенежныхСредств.ПустаяСсылка();
				КонецПопытки;
			КонецЕсли;
		КонецЕсли;
		НовСтр.СтатьяДвиженияДенежныхСредств = СтатьяДДС;
	КонецЦикла;

```

> Если Edit не находит точное совпадение `old_string` из-за CRLF/табов — использовать Python helper (как делали в предыдущих задачах). Helper: прочитать файл, найти границы Шага 8 через regex, заменить блок, записать с CRLF.

- [ ] **Step 4: Commit BSL изменения**

```bash
git add "Documents/РаспределениеФ2/Ext/ObjectModule.bsl"
git commit -m "$(cat <<'EOF'
feat(rashifrovka-unify): Шаг 8 — зеркало ТЧ.Зарплата 1:1 + СтатьяДДС/Направление

Documents/РаспределениеФ2/Ext/ObjectModule.bsl — переписан Шаг 8 заполнения
А_РасшифровкаВыплатыЗарплатаПоФизлицам:

Было: цикл по ТЧ.Распределение Ф2 → ОтработаноЧасов + СуммаНачисления +
Свернуть(...). Связь с ТЧ.Зарплата отсутствовала.

Стало: цикл по уже заполненной ТЧ.Зарплата → ЗаполнитьЗначенияСвойств(...)
с 11 общими полями + НаправлениеДеятельности (из шапки) + СтатьяДДС
(из словаря или Сотр.СтатьяЗарплата с fallback). Свёртка не нужна
(1:1 со строками Зарплата).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 4: Перегенерировать Form.xml подформы

**Files:**
- Create (temp): `_Rarzrabotki/Python/form_def_rashifrovka_v2.json`
- Modify: `Documents/ВедомостьНаВыплатуЗарплатыВКассу/Forms/А_ФормаСпискаРасшифровкиПоФЛ/Ext/Form.xml`

- [ ] **Step 1: Создать обновлённый JSON DSL**

Создать `_Rarzrabotki/Python/form_def_rashifrovka_v2.json`:

```json
{
  "title": "Расшифровка по статье ДДС",
  "properties": {
    "autoTitle": false,
    "windowOpeningMode": "LockOwnerWindow"
  },
  "events": {
    "OnCreateAtServer": "ПриСозданииНаСервере",
    "BeforeClose": "ПередЗакрытием"
  },
  "elements": [
    {
      "table": "А_РасшифровкаВыплатыЗарплатаПоФизлицам",
      "path": "А_РасшифровкаВыплатыЗарплатаПоФизлицам",
      "changeRowSet": true,
      "changeRowOrder": true,
      "on": ["OnStartEdit"],
      "columns": [
        { "input": "ФизическоеЛицо",                "path": "А_РасшифровкаВыплатыЗарплатаПоФизлицам.ФизическоеЛицо",                "markIncomplete": true },
        { "input": "Сотрудник",                     "path": "А_РасшифровкаВыплатыЗарплатаПоФизлицам.Сотрудник" },
        { "input": "Подразделение",                 "path": "А_РасшифровкаВыплатыЗарплатаПоФизлицам.Подразделение" },
        { "input": "ПериодВзаиморасчетов",         "path": "А_РасшифровкаВыплатыЗарплатаПоФизлицам.ПериодВзаиморасчетов" },
        { "input": "СтатьяФинансирования",         "path": "А_РасшифровкаВыплатыЗарплатаПоФизлицам.СтатьяФинансирования" },
        { "input": "СтатьяРасходов",                "path": "А_РасшифровкаВыплатыЗарплатаПоФизлицам.СтатьяРасходов" },
        { "input": "ДокументОснование",            "path": "А_РасшифровкаВыплатыЗарплатаПоФизлицам.ДокументОснование" },
        { "input": "КВыплате",                      "path": "А_РасшифровкаВыплатыЗарплатаПоФизлицам.КВыплате" },
        { "input": "КомпенсацияЗаЗадержкуЗарплаты","path": "А_РасшифровкаВыплатыЗарплатаПоФизлицам.КомпенсацияЗаЗадержкуЗарплаты" },
        { "input": "ГруппаУчетаНачислений",        "path": "А_РасшифровкаВыплатыЗарплатаПоФизлицам.ГруппаУчетаНачислений" },
        { "input": "СтатьяДвиженияДенежныхСредств","path": "А_РасшифровкаВыплатыЗарплатаПоФизлицам.СтатьяДвиженияДенежныхСредств" },
        { "input": "НаправлениеДеятельности",      "path": "А_РасшифровкаВыплатыЗарплатаПоФизлицам.НаправлениеДеятельности" }
      ]
    },
    {
      "group": "horizontal",
      "name": "ГруппаКнопок",
      "children": [
        { "button": "КомандаОК",     "command": "КомандаОК",     "defaultButton": true, "title": "ОК" },
        { "button": "КомандаОтмена", "command": "КомандаОтмена", "title": "Отмена" }
      ]
    }
  ],
  "attributes": [
    {
      "name": "А_РасшифровкаВыплатыЗарплатаПоФизлицам",
      "type": "ValueTable",
      "savedData": true,
      "columns": [
        { "name": "ИдентификаторСтроки",            "type": "UUID" },
        { "name": "Сотрудник",                       "type": "CatalogRef.Сотрудники" },
        { "name": "ФизическоеЛицо",                 "type": "CatalogRef.ФизическиеЛица" },
        { "name": "Подразделение",                  "type": "CatalogRef.ПодразделенияОрганизаций" },
        { "name": "ПериодВзаиморасчетов",          "type": "dateTime" },
        { "name": "СтатьяФинансирования",          "type": "CatalogRef.СтатьиФинансированияЗарплата" },
        { "name": "СтатьяРасходов",                 "type": "CatalogRef.СтатьиРасходовЗарплата" },
        { "name": "ДокументОснование",             "type": "DocumentRef.НачислениеЗарплаты" },
        { "name": "КВыплате",                       "type": "decimal(15,2)" },
        { "name": "КомпенсацияЗаЗадержкуЗарплаты", "type": "decimal(15,2)" },
        { "name": "ГруппаУчетаНачислений",         "type": "CatalogRef.ГруппыУчетаНачисленийИУдержаний" },
        { "name": "СтатьяДвиженияДенежныхСредств", "type": "CatalogRef.СтатьиДвиженияДенежныхСредств" },
        { "name": "НаправлениеДеятельности",       "type": "CatalogRef.НаправленияДеятельности" }
      ]
    },
    { "name": "ФизическоеЛицо", "type": "CatalogRef.ФизическиеЛица" },
    { "name": "Организация",    "type": "CatalogRef.Организации" }
  ],
  "commands": [
    { "name": "КомандаОК",     "title": "ОК",     "action": "ОК" },
    { "name": "КомандаОтмена", "title": "Отмена", "action": "Отмена" }
  ]
}
```

> Примечание: `ДокументОснование` в JSON DSL поставлен как `DocumentRef.НачислениеЗарплаты` (один тип) — для UI этого достаточно: пользователь будет видеть значение из ТЧ. Composite-тип (15 типов) определяется уровне ТЧ документа (Documents/ВКассу.xml), а атрибут формы — это копия ТЗ, в которой одна из колонок имеет тип-ссылка. По факту в Расшифровке ДокументОснование заполнится из ТЧ.Зарплата.ДокументОснование (один из 15 типов).
> Если при загрузке возникнет ошибка типа — поправить тип в JSON на `Any` или составной (`DocumentRef.НачислениеЗарплаты | DocumentRef.ОплатаПоСреднемуЗаработку | ...`).

- [ ] **Step 2: Запустить /form-compile**

Run:
```powershell
& "C:\Configuration_downloads\BASERP25\.claude\skills\form-compile\scripts\form-compile.ps1" -JsonPath "_Rarzrabotki\Python\form_def_rashifrovka_v2.json" -OutputPath "Documents\ВедомостьНаВыплатуЗарплатыВКассу\Forms\А_ФормаСпискаРасшифровкиПоФЛ\Ext\Form.xml"
```

Expected: `[OK] Compiled: ... Form.xml`.

Если warning про `UUID` — после компиляции править вручную: найти `<v8:Type>UUID</v8:Type>` для колонки ИдентификаторСтроки и заменить на `<v8:Type>v8:UUID</v8:Type>`.

- [ ] **Step 3: Удалить JSON helper**

```bash
rm "_Rarzrabotki/Python/form_def_rashifrovka_v2.json"
```

- [ ] **Step 4: Валидация формы**

Run:
```powershell
& "C:\Configuration_downloads\BASERP25\.claude\skills\form-validate\scripts\form-validate.ps1" -FormPath "Documents\ВедомостьНаВыплатуЗарплатыВКассу\Forms\А_ФормаСпискаРасшифровкиПоФЛ\Ext\Form.xml"
```

Expected: `0 errors` (warnings про version 2.13 — игнорируем).

- [ ] **Step 5: Commit формы**

```bash
git add "Documents/ВедомостьНаВыплатуЗарплатыВКассу/Forms/А_ФормаСпискаРасшифровкиПоФЛ/Ext/Form.xml"
git commit -m "$(cat <<'EOF'
feat(rashifrovka-unify): форма А_ФормаСпискаРасшифровкиПоФЛ — 13 колонок

Form.xml перегенерирован через /form-compile с новым набором колонок
(13 вместо 8). Колонки соответствуют новой структуре ТЧ:
- удалена ОтработаноЧасов
- переименована Сумма → КВыплате
- добавлены 6 новых: ПериодВзаиморасчетов, СтатьяФинансирования,
  СтатьяРасходов, ДокументОснование, КомпенсацияЗаЗадержкуЗарплаты,
  ГруппаУчетаНачислений
- наши специфичные сохранены: НаправлениеДеятельности, СтатьяДДС

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 5: Доставка через db-load-xml + db-update

**Files:**
- Deliver: `Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml`, `Forms/А_ФормаСпискаРасшифровкиПоФЛ.xml`, `Forms/А_ФормаСпискаРасшифровкиПоФЛ/Ext/Form.xml`, `Documents/РаспределениеФ2/Ext/ObjectModule.bsl`

- [ ] **Step 1: Проверить что Designer закрыт**

Спросить пользователя:
> Designer закрыт? Готов к доставке (структурное изменение ТЧ)?

Дождаться подтверждения.

- [ ] **Step 2: Partial load**

Run:
```powershell
& "C:\Configuration_downloads\BASERP25\.claude\skills\db-load-xml\scripts\db-load-xml.ps1" -V8Path "C:\Program Files\1cv8\8.3.20.1914\bin" -InfoBaseServer "SQLSERVER" -InfoBaseRef "BaseERP" -UserName "Администратор" -Password "24043" -ConfigDir "C:\Configuration_downloads\BASERP25" -Mode Partial -Files "Documents/ВедомостьНаВыплатуЗарплатыВКассу.xml,Documents/ВедомостьНаВыплатуЗарплатыВКассу/Forms/А_ФормаСпискаРасшифровкиПоФЛ.xml,Documents/РаспределениеФ2/Ext/ObjectModule.bsl"
```

Expected: `Load completed successfully`.

Если упадёт с ошибкой XML — диагностировать (нет валидного DocumentRef в ДокументОснование, нет UUID и т.п.).

- [ ] **Step 3: Dynamic update**

Run:
```powershell
& "C:\Configuration_downloads\BASERP25\.claude\skills\db-update\scripts\db-update.ps1" -V8Path "C:\Program Files\1cv8\8.3.20.1914\bin" -InfoBaseServer "SQLSERVER" -InfoBaseRef "BaseERP" -UserName "Администратор" -Password "24043" -Dynamic "+"
```

Expected: `Обновление конфигурации успешно завершено`.

Если упадёт с реструктуризацией → попробовать без `-Dynamic+`:
```powershell
& "C:\Configuration_downloads\BASERP25\.claude\skills\db-update\scripts\db-update.ps1" -V8Path "C:\Program Files\1cv8\8.3.20.1914\bin" -InfoBaseServer "SQLSERVER" -InfoBaseRef "BaseERP" -UserName "Администратор" -Password "24043"
```

Если требует перезапуск сервера — спросить пользователя.

---

## Task 6: Регрессия + проверка миграции данных

**Files:** не правим — только запускаем тесты + диагностика.

- [ ] **Step 1: Smoke метаданных формы**

Run:
```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_vkassu_subform_metadata.py"
```

Expected: тест может УПАСТЬ — он проверяет старый список реквизитов (ОтработаноЧасов, Сумма). Это **ожидаемо**, исправим в Task 7.

- [ ] **Step 2: Проверка миграции данных (Сумма → КВыплате)**

Запустить однократный Python COM:
```python
"""Smoke: проверить что после переименования Сумма→КВыплате значения сохранились в существующих ВКассу."""
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import win32com.client
v8 = win32com.client.Dispatch("V83.COMConnector")
erp = v8.Connect('Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"')

q = erp.NewObject("Запрос")
q.Текст = '''ВЫБРАТЬ ПЕРВЫЕ 5
    Вед.Номер КАК Номер,
    СУММА(Р.КВыплате) КАК ΣКВыплате,
    КОЛИЧЕСТВО(*) КАК Кол
ИЗ Документ.ВедомостьНаВыплатуЗарплатыВКассу.А_РасшифровкаВыплатыЗарплатаПоФизлицам КАК Р
    ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.ВедомостьНаВыплатуЗарплатыВКассу КАК Вед ПО Р.Ссылка = Вед.Ссылка
ГДЕ НЕ Вед.ПометкаУдаления
СГРУППИРОВАТЬ ПО Вед.Номер
УПОРЯДОЧИТЬ ПО Вед.Номер'''
rs = q.Выполнить()
sel = rs.Выбрать()
print("ВКассу | строк | Σ КВыплате")
while sel.Следующий():
    print(f"  {sel.Номер}: {int(sel.Кол)} строк, Σ={float(sel.ΣКВыплате):,.2f}")
```

Ожидаем: КВыплате в существующих ВКассу содержит миграционные значения (равные старой Сумме). Если 0 — значит миграция UUID не сработала, нужно перепровести Ф2.

- [ ] **Step 3: Регрессия Ф2-smoke (НЕ должно сломаться)**

Run:
```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_create_vedomost_vkassu_iz_f2_smoke.py"
```

Expected: PASS (smoke не лезет в А_Расшифровку, проверяет ТЧ.Зарплата и шапку).

Это перепроведёт Ф2 №000000026 → пересоздаст ВКассу с новой структурой А_Расшифровки (31 строка с заполненными новыми полями).

- [ ] **Step 4: Smoke ОтражениеЗП (тоже не должно сломаться)**

Run:
```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_otrazhenie_vkassu_smoke.py"
```

Expected: PASS (создаёт 7 ВКассу с обновлёнными А_Расшифровками).

---

## Task 7: Обновить тесты под новую структуру

**Files:**
- Modify: `_Rarzrabotki/Python/test/test_create_vedomost_vkassu_iz_f2_rashifrovka.py`
- Modify: `_Rarzrabotki/Python/test/test_vkassu_subform_returns_changes.py`
- Modify: `_Rarzrabotki/Python/test/test_vkassu_subform_metadata.py`

- [ ] **Step 1: Поправить `test_create_vedomost_vkassu_iz_f2_rashifrovka.py`**

Заменить блок проверки ОтработаноЧасов:

Найти:
```python
    # Σ ОтработаноЧасов в А_Расшифровке должна быть = Σ Распределение.ОтработаноЧасов
    sum_h_f2 = sum(float(f2.Распределение.Получить(i).ОтработаноЧасов)
                   for i in range(f2.Распределение.Количество()))
    sum_h_vk = sum(float(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Получить(i).ОтработаноЧасов)
                   for i in range(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Количество()))
    print(f"Σ ОтработаноЧасов: Ф2={sum_h_f2}, ВКассу={sum_h_vk}")
    if abs(sum_h_f2 - sum_h_vk) > 0.01:
        print(f"FAIL: Несовпадение часов: {sum_h_f2} != {sum_h_vk}")
        sys.exit(1)
```

Удалить (часов в А_Расшифровке больше нет).

Найти:
```python
    sum_s_vk = sum(float(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Получить(i).Сумма)
                   for i in range(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Количество()))
```

Заменить на:
```python
    sum_s_vk = sum(float(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Получить(i).КВыплате)
                   for i in range(vk.А_РасшифровкаВыплатыЗарплатаПоФизлицам.Количество()))
```

Также строка вывода:
```python
    print(f"Σ Сумма: Ф2={sum_s_f2:,.2f}, ВКассу={sum_s_vk:,.2f}")
```

Заменить:
```python
    print(f"Σ КВыплате: Ф2={sum_s_f2:,.2f}, ВКассу={sum_s_vk:,.2f}")
```

Финальный print:
```python
    print(f"\nPASS: rashifrovka test пройден ({cnt} строк)")
```
Оставить.

- [ ] **Step 2: Запустить rashifrovka-тест**

Run:
```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_create_vedomost_vkassu_iz_f2_rashifrovka.py"
```

Expected: PASS (Σ КВыплате в А_Расшифровке = 348 800 для Ф2 №000000026, столько же что и в Ф2.Распределение.СуммаНачисления).

- [ ] **Step 3: Поправить `test_vkassu_subform_returns_changes.py`**

В тесте сохранения данных по `Сумма` и `Часы`:

Найти все упоминания `.Сумма` (поля А_Расшифровки) → заменить на `.КВыплате`.

Найти и удалить упоминания `.ОтработаноЧасов` в копировании данных (структура `исходные` имеет ключ `"Часы"` — оставить только если используется; иначе удалить).

Точные правки:
- В `исходные.append({...})` убрать `"Часы": float(r.ОтработаноЧасов or 0)`
- В цикле копирования полей при создании новой строки убрать `нс.ОтработаноЧасов = ис["Часы"]`
- Все `.Сумма` → `.КВыплате`

- [ ] **Step 4: Запустить subform_returns_changes**

Run:
```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_vkassu_subform_returns_changes.py"
```

Expected: PASS.

- [ ] **Step 5: Поправить `test_vkassu_subform_metadata.py`**

Найти:
```python
required_attrs = ["ФизическоеЛицо", "Сотрудник", "Подразделение",
                  "СтатьяДвиженияДенежныхСредств", "НаправлениеДеятельности",
                  "ОтработаноЧасов", "Сумма", "ИдентификаторСтроки"]
```

Заменить на:
```python
required_attrs = ["ИдентификаторСтроки", "Сотрудник", "ФизическоеЛицо", "Подразделение",
                  "ПериодВзаиморасчетов", "СтатьяФинансирования", "СтатьяРасходов",
                  "ДокументОснование", "КВыплате", "КомпенсацияЗаЗадержкуЗарплаты",
                  "ГруппаУчетаНачислений", "НаправлениеДеятельности",
                  "СтатьяДвиженияДенежныхСредств"]
```

И в выводе:
```python
print(f"OK: ТЧ имеет все 8 нужных реквизитов")
```
→
```python
print(f"OK: ТЧ имеет все {len(required_attrs)} нужных реквизитов")
```

- [ ] **Step 6: Запустить metadata-тест**

Run:
```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_vkassu_subform_metadata.py"
```

Expected: PASS (13 реквизитов).

- [ ] **Step 7: Commit тестов**

```bash
git add "_Rarzrabotki/Python/test/test_create_vedomost_vkassu_iz_f2_rashifrovka.py" "_Rarzrabotki/Python/test/test_vkassu_subform_returns_changes.py" "_Rarzrabotki/Python/test/test_vkassu_subform_metadata.py"
git commit -m "$(cat <<'EOF'
test(rashifrovka-unify): обновлены 3 теста под новую структуру А_Расшифровки

- test_create_vedomost_vkassu_iz_f2_rashifrovka: убрана проверка ОтработаноЧасов,
  переименовано Сумма → КВыплате (Σ КВыплате в А_Расшифровке = Σ Ф2.Распределение)
- test_vkassu_subform_returns_changes: Сумма → КВыплате, убраны Часы
  в копировании полей
- test_vkassu_subform_metadata: required_attrs теперь 13 реквизитов
  (зеркало ТЧ.Зарплата 11 + 2 наших)

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 8: Финал — регрессия всех тестов + knowledge

**Files:** не правим — только запускаем + knowledge.

- [ ] **Step 1: Полная регрессия — 8 тестов**

Run (последовательно — паралельно через `&&` MCP не любит):
```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_create_vedomost_vkassu_iz_f2_smoke.py"
```
Expected: PASS.

```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_create_vedomost_vkassu_iz_f2_idempotency.py"
```
Expected: PASS.

```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_create_vedomost_vkassu_iz_f2_rashifrovka.py"
```
Expected: PASS.

```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_otrazhenie_vkassu_smoke.py"
```
Expected: PASS (7 ВКассу).

```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_otrazhenie_vkassu_idempotency.py"
```
Expected: PASS.

```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_otrazhenie_vkassu_statyaddc.py"
```
Expected: PASS.

```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_vkassu_subform_metadata.py"
```
Expected: PASS.

```
C:\Python313\python.exe "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Python\test\test_vkassu_subform_returns_changes.py"
```
Expected: PASS.

Если хоть один FAIL — диагностировать.

- [ ] **Step 2: Обновить knowledge 11 (Урок №7 → СтатьяДДС + ГруппаУчётаНачислений из Зарплата)**

В `_Rarzrabotki/notebook/knowledge_А_ОтражениеЗПпоКазне/11_vedomost_iz_raspredeleniyaF2.md`:

В разделе "3. Критичные находки и уроки" добавить НОВЫЙ урок (после Урока №7):

```markdown
### Урок №8: ТЧ А_Расшифровка зеркалит ТЧ Зарплата 1:1 (унификация 2026-05-27)

С 2026-05-27 ТЧ `А_РасшифровкаВыплатыЗарплатаПоФизлицам` имеет 13 полей:
11 общих с ТЧ.Зарплата (ИдентификаторСтроки, Сотрудник, ФизЛицо, Подразделение,
ПериодВзаиморасчетов, СтатьяФинансирования, СтатьяРасходов, ДокументОснование,
КВыплате, КомпенсацияЗаЗадержкуЗарплаты, ГруппаУчётаНачислений) + 2 наших
(НаправлениеДеятельности, СтатьяДвиженияДенежныхСредств).

В Шаге 8 заполняется через `ЗаполнитьЗначенияСвойств(НовСтр, СтрЗ, ПоляЗеркала)`
по уже заполненной ТЧ.Зарплата — связь 1:1 через ИдентификаторСтроки.
Свёртка А_Расшифровки больше не нужна.

Spec/Plan:
- [`docs/superpowers/specs/2026-05-27-rashifrovka-unify-with-zarplata-design.md`](../../../docs/superpowers/specs/2026-05-27-rashifrovka-unify-with-zarplata-design.md)
- [`docs/superpowers/plans/2026-05-27-rashifrovka-unify-with-zarplata-plan.md`](../../../docs/superpowers/plans/2026-05-27-rashifrovka-unify-with-zarplata-plan.md)
```

- [ ] **Step 3: Обновить knowledge 4. Эталонный документ — 13 полей**

В `11_vedomost_iz_raspredeleniyaF2.md` найти таблицу «4. Эталонный документ» → найти строку:
```
| ВКассу Зарплата строк | **31** ... |
| ВКассу Зарплата.ГруппаУчётаНачислений | "Зарплата (661)" во всех строках |
```

После неё добавить:
```
| ВКассу А_Расшифровка структура | **13 полей** (с 2026-05-27 — зеркало ТЧ.Зарплата + НаправлениеДеятельности + СтатьяДДС) |
```

- [ ] **Step 4: Commit knowledge**

```bash
git add "_Rarzrabotki/notebook/knowledge_А_ОтражениеЗПпоКазне/11_vedomost_iz_raspredeleniyaF2.md"
git commit -m "$(cat <<'EOF'
docs(knowledge): Урок №8 — ТЧ А_Расшифровка зеркало ТЧ.Зарплата (унификация)

С 2026-05-27 ТЧ А_РасшифровкаВыплатыЗарплатаПоФизлицам имеет 13 полей:
- 11 зеркалят ТЧ.Зарплата (включая ПериодВзаиморасчетов, ГруппаУчёта,
  ДокументОснование composite ref на 15 типов)
- 2 наших специфичных (НаправлениеДеятельности, СтатьяДДС)

Шаг 8 в РаспределениеФ2.СоздатьВедомостьНаВыплатуВКассу теперь использует
ЗаполнитьЗначенияСвойств(НовСтр, СтрЗ, ПоляЗеркала) — связь 1:1 со строками
Зарплата через ИдентификаторСтроки. Свёртка не нужна.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

- [ ] **Step 5: Финальный отчёт**

Сообщить пользователю:
- 8 тестов PASS
- Структура ТЧ А_Расшифровка унифицирована с ТЧ.Зарплата (13 полей)
- Форма А_ФормаСпискаРасшифровкиПоФЛ перегенерирована с 13 колонками
- Существующие ВКассу обновлены (через перепроведение Ф2 в smoke-тесте)
- Жду «push» от пользователя для отправки в `origin/main`

---

## Self-Review (выполнено)

**1. Spec coverage:**
- §2 «Изменения структуры» → Task 2 (XML)
- §3.1 «Documents/ВКассу.xml» → Task 2
- §3.2 «РаспределениеФ2.Шаг 8» → Task 3
- §3.3 «А_ОтражениеЗПпоКазне» → не меняем (отмечено в плане)
- §3.4 «Форма» → Task 4
- §3.5 «Тесты» → Task 7
- §4 «Доставка» → Task 5 (с fallback на полный update)
- §5 «Существующие данные» → Task 6 (проверка миграции UUID)
- §6 «Out-of-scope» → не реализуется
- §7 «Acceptance критерии» → проверяются через 8 тестов в Task 8

**2. Placeholder scan:** все шаги содержат конкретный код / команды / regex. UUIDs обозначены как «подставить из Task 1 Step 4» — это шаг плана, не placeholder.

**3. Type consistency:**
- Имя поля `КВыплате` (не «К_выплате», не «Выплата»): везде одинаковое (BSL/JSON/тесты)
- Имя ТЧ `А_РасшифровкаВыплатыЗарплатаПоФизлицам`: везде идентичное (15 упоминаний)
- Имя 6 новых полей: `ПериодВзаиморасчетов`, `СтатьяФинансирования`, `СтатьяРасходов`, `ДокументОснование`, `КомпенсацияЗаЗадержкуЗарплаты`, `ГруппаУчетаНачислений` — везде одинаковые
- Сигнатура `Ф2.СоздатьВедомостьНаВыплатуВКассуПоРаспределениюФ2(СтатьиДДСПоКлючу, СсылкаНаОтражениеЗП)` не меняется

Плану — зелёный свет.

---

## Execution Handoff

**Plan complete and saved to `docs/superpowers/plans/2026-05-27-rashifrovka-unify-with-zarplata-plan.md`. Two execution options:**

**1. Subagent-Driven (recommended)** — диспатчим свежего subagent на каждую задачу, ревью между задачами, быстрая итерация. **REQUIRED SUB-SKILL:** `superpowers:subagent-driven-development`.

**2. Inline Execution** — выполняем задачи в этой сессии через `superpowers:executing-plans`, batch с чекпоинтами для ревью.

**Which approach?**
