# А_ОбработкаДисбалансаПоСтатьямБаланса Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Доработать внешний отчёт `А_ОбработкаДисбалансаПоСтатьямБаланса.erf`: добавить кнопки «Анализ расхождений» (FULL OUTER ПАП vs РСКПС/РСППС через `ДокументРегистратор`, заполняет ТЧ виновниками) и «Перепровести» (per-документ транзакция). E2E тестирование через Python COM.

**Architecture:** ExternalReport (каркас уже есть), доп. реквизиты + ТЧ `ДокументыРасхождения` + ObjectModule.bsl с 3 процедурами + Form/Module.bsl с 2 обработчиками. Запускается на сервере 1С (через UI или Python COM вызовом `ВнешниеОтчеты.Создать()`).

**Tech Stack:** 1С:Предприятие 8.3, BSL (BAS ERP 2.5), Python 3.11 + `win32com.client` (V83.COMConnector), skill `/erf-build` для сборки ERF.

**Spec:** [docs/superpowers/specs/2026-05-25-obrabotka-disbalansa-statyam-design.md](../specs/2026-05-25-obrabotka-disbalansa-statyam-design.md)

---

## Pre-execution checklist

- [ ] **Workdir:** `C:\Configuration_downloads\BASERP25` (основная конфигурация)
- [ ] **DB:** `Srvr="localhost";Ref="BaseERP";Usr="Администратор";Pwd="24043"` — доступна
- [ ] **Изучить spec** разделы 3-5 (структура реквизитов, алгоритмы) перед началом
- [ ] **Изучить образец** `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоПодразделениям/Ext/ObjectModule.bsl` (паттерн АнализДокументов/СоздатьДокументы)

---

### Task 1: Baseline snapshot — зафиксировать текущий стан 19 виновников

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_baseline.py`
- Update: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/_artifacts/obrabotka_baseline.json`

Acceptance T9 (Test 1) проверяет что обработка находит **точно** те же 19 виновников. Сначала зафиксируем их машиночитаемо.

- [ ] **Step 1: Создать Python скрипт baseline**

Create file `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_baseline.py`:

```python
# -*- coding: utf-8 -*-
"""
Baseline для обработки А_ОбработкаДисбалансаПоСтатьямБаланса.

Запускаем перед реализацией — фиксирует ровно те 19 виновников апреля 2026
по 9 подразделениям статья ЗПП, Source=РСППС, которые ручной анализ нашёл
(Σ Δ по подразделениям совпадает с плугами Свода 00000000004 до копейки).

Результат → _artifacts/obrabotka_baseline.json — используется как oracle
в Test 1 (obrabotka_test1_analiz.py).
"""
import sys, io, json, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp, ARTIFACTS_DIR

erp = connect_erp()

q = erp.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ
    Т.Регистратор КАК Док,
    ПРЕДСТАВЛЕНИЕ(Т.Регистратор) КАК ДокИмя,
    Т.Подразделение.Код КАК ПодрКод,
    ПРЕДСТАВЛЕНИЕ(Т.Подразделение) КАК ПодрИмя,
    СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                ТОГДА Т.Сумма ИНАЧЕ -Т.Сумма КОНЕЦ) КАК ПАП_Sign
ПОМЕСТИТЬ втПАП
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
ГДЕ Т.Период МЕЖДУ ДАТАВРЕМЯ(2026,4,1) И ДАТАВРЕМЯ(2026,4,30,23,59,59)
    И Т.Организация.КодПоЕДРПОУ = "40645273"
    И Т.Подразделение.Код В ("00-001029","00-001022","00-001026","0Ц-000004","00-001003","0Ц-000001","00-000210","00-000219","00-000190")
    И Т.Источник = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСПоставщикамиПоСрокам)
    И Т.Статья = ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьПередПоставщиками)
СГРУППИРОВАТЬ ПО Т.Регистратор, ПРЕДСТАВЛЕНИЕ(Т.Регистратор), Т.Подразделение.Код, ПРЕДСТАВЛЕНИЕ(Т.Подразделение)
;
ВЫБРАТЬ
    Р.ДокументРегистратор КАК Док,
    ПРЕДСТАВЛЕНИЕ(Р.ДокументРегистратор) КАК ДокИмя,
    ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
        ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
        ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
    КОНЕЦ КАК ПодрКод,
    СУММА(ВЫБОР КОГДА Р.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                ТОГДА Р.ДолгУпр ИНАЧЕ -Р.ДолгУпр КОНЕЦ) КАК РСППС_DolgSign
ПОМЕСТИТЬ втРСППС
ИЗ РегистрНакопления.РасчетыСПоставщикамиПоСрокам КАК Р
    ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
        ПО Р.АналитикаУчетаПоПартнерам = АП.КлючАналитики
ГДЕ Р.Период МЕЖДУ ДАТАВРЕМЯ(2026,4,1) И ДАТАВРЕМЯ(2026,4,30,23,59,59)
    И АП.Организация.КодПоЕДРПОУ = "40645273"
    И АП.Партнер <> ЗНАЧЕНИЕ(Справочник.Партнеры.НашеПредприятие)
    И ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
        ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
        ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
    КОНЕЦ В ("00-001029","00-001022","00-001026","0Ц-000004","00-001003","0Ц-000001","00-000210","00-000219","00-000190")
СГРУППИРОВАТЬ ПО Р.ДокументРегистратор, ПРЕДСТАВЛЕНИЕ(Р.ДокументРегистратор),
    ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
        ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
        ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
    КОНЕЦ
;
ВЫБРАТЬ
    ЕСТЬNULL(ПАП.ПодрКод, РСППС.ПодрКод) КАК ПодрКод,
    ЕСТЬNULL(ПАП.ДокИмя, РСППС.ДокИмя) КАК ДокИмя,
    ЕСТЬNULL(ПАП.ПАП_Sign, 0) КАК ПАП_Sign,
    -ЕСТЬNULL(РСППС.РСППС_DolgSign, 0) КАК РСППС_Inv,
    ЕСТЬNULL(ПАП.ПАП_Sign, 0) + ЕСТЬNULL(РСППС.РСППС_DolgSign, 0) КАК Дельта
ИЗ втПАП КАК ПАП
    ПОЛНОЕ СОЕДИНЕНИЕ втРСППС КАК РСППС
    ПО ПАП.Док = РСППС.Док И ПАП.ПодрКод = РСППС.ПодрКод
ГДЕ ЕСТЬNULL(ПАП.ПАП_Sign, 0) + ЕСТЬNULL(РСППС.РСППС_DolgSign, 0) > 0.01
   ИЛИ ЕСТЬNULL(ПАП.ПАП_Sign, 0) + ЕСТЬNULL(РСППС.РСППС_DolgSign, 0) < -0.01
УПОРЯДОЧИТЬ ПО ПодрКод, Дельта
"""
r = q.Выполнить().Выгрузить()
rows = []
sums = {}
for i in range(r.Количество()):
    rec = r.Получить(i)
    code = str(rec.ПодрКод)
    delta = float(rec.Дельта)
    sums[code] = sums.get(code, 0) + delta
    rows.append({
        "ПодрКод": code,
        "ДокИмя": str(rec.ДокИмя),
        "ПАП_Sign": float(rec.ПАП_Sign),
        "РСППС_Inv": float(rec.РСППС_Inv),
        "Дельта": delta,
    })

baseline = {"rows": rows, "sums_by_podr": sums, "total_rows": len(rows)}
out = os.path.join(ARTIFACTS_DIR, "obrabotka_baseline.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(baseline, f, ensure_ascii=False, indent=2)

print(f"Baseline зафиксирован: {len(rows)} строк, {len(sums)} подразделений")
print(f"Артефакт: {out}")
for code, s in sorted(sums.items()):
    print(f"  {code}: Σ Δ = {s:+,.2f}")
```

- [ ] **Step 2: Запустить baseline**

Run: `python _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_baseline.py`

Expected output:
```
Baseline зафиксирован: 19 строк, 9 подразделений
  00-000190: Σ Δ = +45,005.00
  00-000210: Σ Δ = +25,880.92
  00-000219: Σ Δ = +30,710.71
  00-001003: Σ Δ = -13,359.32
  00-001022: Σ Δ = -156,115.38
  00-001026: Σ Δ = -39,936.00
  00-001029: Σ Δ = -1,788,600.00
  0Ц-000001: Σ Δ = -780.00
  0Ц-000004: Σ Δ = -37,800.00
```

⚠️ Если число строк ≠ 19 или суммы по подр не совпадают с этими — STOP, разобраться (возможно пользователь успел ещё что-то перепровести, baseline нужно зафиксировать актуальный).

- [ ] **Step 3: Commit baseline**

```bash
git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_baseline.py \
        _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/_artifacts/obrabotka_baseline.json
git commit -m "$(cat <<'EOF'
test(balans_klient): baseline для obrabotka_disbalansa_statyam — 19 виновников апр26

Зафиксировал текущий снимок 19 первичных документов с расхождением
ПАП vs РСППС по статье ЗПП для 9 подразделений за апрель 2026.
Используется как oracle в Test 1 после реализации обработки.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 2: XML обработки — добавить реквизиты, ТЧ, команды

**Files:**
- Modify: `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса.xml`

Добавляем к существующим 2 реквизитам ещё 2 (`Подразделение`, `ПоказыватьВсе`) + ТЧ `ДокументыРасхождения` (10 колонок) + 2 Команды (`АнализРасхождений`, `Перепровести`).

Образец — `А_ОбработкаДисбалансаПоПодразделениям.xml` (структура реквизитов 78-186, ТЧ 187-586).

- [ ] **Step 1: Прочитать образец как референс для XML структур**

Run: `head -200 _Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоПодразделениям.xml`

Ничего не делать с output — запомнить структуру блоков `<Attribute>` и `<TabularSection>`.

- [ ] **Step 2: Открыть текущий файл и спланировать вставку**

Read: `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса.xml`

Точки вставки:
- После закрывающего `</Attribute>` второго реквизита (`ОкончаниеПериода`, строка ~118) и **перед** `<Form>Форма</Form>` (строка ~119)

- [ ] **Step 3: Вставить 2 реквизита (Подразделение + ПоказыватьВсе) + ТЧ + 2 Команды**

Use Edit tool на файле `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса.xml`.

`old_string`:
```xml
			<Form>Форма</Form>
		</ChildObjects>
	</ExternalReport>
</MetaDataObject>
```

`new_string` — большая вставка (~250 строк):

```xml
			<Attribute uuid="b1000001-0000-0000-0000-000000000001">
				<Properties>
					<Name>Подразделение</Name>
					<Synonym>
						<v8:item>
							<v8:lang>ru</v8:lang>
							<v8:content>Подразделение</v8:content>
						</v8:item>
					</Synonym>
					<Comment/>
					<Type>
						<v8:Type>cfg:CatalogRef.СтруктураПредприятия</v8:Type>
					</Type>
					<PasswordMode>false</PasswordMode>
					<Format/>
					<EditFormat/>
					<ToolTip/>
					<MarkNegatives>false</MarkNegatives>
					<Mask/>
					<MultiLine>false</MultiLine>
					<ExtendedEdit>false</ExtendedEdit>
					<MinValue xsi:nil="true"/>
					<MaxValue xsi:nil="true"/>
					<FillChecking>DontCheck</FillChecking>
					<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
					<ChoiceParameterLinks/>
					<ChoiceParameters/>
					<QuickChoice>Auto</QuickChoice>
					<CreateOnInput>Auto</CreateOnInput>
					<ChoiceForm/>
					<LinkByType/>
					<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
				</Properties>
			</Attribute>
			<Attribute uuid="b1000002-0000-0000-0000-000000000001">
				<Properties>
					<Name>ПоказыватьВсе</Name>
					<Synonym>
						<v8:item>
							<v8:lang>ru</v8:lang>
							<v8:content>Показывать все (включая обработанные)</v8:content>
						</v8:item>
					</Synonym>
					<Comment/>
					<Type>
						<v8:Type>xs:boolean</v8:Type>
					</Type>
					<PasswordMode>false</PasswordMode>
					<Format/>
					<EditFormat/>
					<ToolTip/>
					<MarkNegatives>false</MarkNegatives>
					<Mask/>
					<MultiLine>false</MultiLine>
					<ExtendedEdit>false</ExtendedEdit>
					<MinValue xsi:nil="true"/>
					<MaxValue xsi:nil="true"/>
					<FillChecking>DontCheck</FillChecking>
					<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
					<ChoiceParameterLinks/>
					<ChoiceParameters/>
					<QuickChoice>Auto</QuickChoice>
					<CreateOnInput>Auto</CreateOnInput>
					<ChoiceForm/>
					<LinkByType/>
					<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
				</Properties>
			</Attribute>
			<TabularSection uuid="b1000010-0000-0000-0000-000000000001">
				<InternalInfo>
					<xr:GeneratedType name="ReportTabularSection.А_ОбработкаДисбалансаПоСтатьямБаланса.ДокументыРасхождения" category="TabularSection">
						<xr:TypeId>b1000011-0000-0000-0000-000000000001</xr:TypeId>
						<xr:ValueId>b1000012-0000-0000-0000-000000000001</xr:ValueId>
					</xr:GeneratedType>
					<xr:GeneratedType name="ReportTabularSectionRow.А_ОбработкаДисбалансаПоСтатьямБаланса.ДокументыРасхождения" category="TabularSectionRow">
						<xr:TypeId>b1000013-0000-0000-0000-000000000001</xr:TypeId>
						<xr:ValueId>b1000014-0000-0000-0000-000000000001</xr:ValueId>
					</xr:GeneratedType>
				</InternalInfo>
				<Properties>
					<Name>ДокументыРасхождения</Name>
					<Synonym>
						<v8:item>
							<v8:lang>ru</v8:lang>
							<v8:content>Документы расхождения</v8:content>
						</v8:item>
					</Synonym>
					<Comment/>
					<ToolTip/>
					<FillChecking>DontCheck</FillChecking>
				</Properties>
				<ChildObjects>
					<Attribute uuid="b1000020-0000-0000-0000-000000000001">
						<Properties>
							<Name>Обработан</Name>
							<Synonym><v8:item><v8:lang>ru</v8:lang><v8:content>Обработан</v8:content></v8:item></Synonym>
							<Comment/>
							<Type><v8:Type>xs:boolean</v8:Type></Type>
							<PasswordMode>false</PasswordMode><Format/><EditFormat/><ToolTip/>
							<MarkNegatives>false</MarkNegatives><Mask/><MultiLine>false</MultiLine>
							<ExtendedEdit>false</ExtendedEdit><MinValue xsi:nil="true"/><MaxValue xsi:nil="true"/>
							<FillFromFillingValue>false</FillFromFillingValue><FillValue xsi:nil="true"/>
							<FillChecking>DontCheck</FillChecking><ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
							<ChoiceParameterLinks/><ChoiceParameters/><QuickChoice>Auto</QuickChoice>
							<CreateOnInput>Auto</CreateOnInput><ChoiceForm/><LinkByType/>
							<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
						</Properties>
					</Attribute>
					<Attribute uuid="b1000021-0000-0000-0000-000000000001">
						<Properties>
							<Name>Документ</Name>
							<Synonym><v8:item><v8:lang>ru</v8:lang><v8:content>Документ</v8:content></v8:item></Synonym>
							<Comment/>
							<Type><v8:TypeSet>cfg:DocumentRef</v8:TypeSet></Type>
							<PasswordMode>false</PasswordMode><Format/><EditFormat/><ToolTip/>
							<MarkNegatives>false</MarkNegatives><Mask/><MultiLine>false</MultiLine>
							<ExtendedEdit>false</ExtendedEdit><MinValue xsi:nil="true"/><MaxValue xsi:nil="true"/>
							<FillFromFillingValue>false</FillFromFillingValue><FillValue xsi:nil="true"/>
							<FillChecking>DontCheck</FillChecking><ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
							<ChoiceParameterLinks/><ChoiceParameters/><QuickChoice>Auto</QuickChoice>
							<CreateOnInput>Auto</CreateOnInput><ChoiceForm/><LinkByType/>
							<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
						</Properties>
					</Attribute>
					<Attribute uuid="b1000022-0000-0000-0000-000000000001">
						<Properties>
							<Name>Подразделение</Name>
							<Synonym><v8:item><v8:lang>ru</v8:lang><v8:content>Подразделение</v8:content></v8:item></Synonym>
							<Comment/>
							<Type><v8:Type>cfg:CatalogRef.СтруктураПредприятия</v8:Type></Type>
							<PasswordMode>false</PasswordMode><Format/><EditFormat/><ToolTip/>
							<MarkNegatives>false</MarkNegatives><Mask/><MultiLine>false</MultiLine>
							<ExtendedEdit>false</ExtendedEdit><MinValue xsi:nil="true"/><MaxValue xsi:nil="true"/>
							<FillFromFillingValue>false</FillFromFillingValue><FillValue xsi:nil="true"/>
							<FillChecking>DontCheck</FillChecking><ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
							<ChoiceParameterLinks/><ChoiceParameters/><QuickChoice>Auto</QuickChoice>
							<CreateOnInput>Auto</CreateOnInput><ChoiceForm/><LinkByType/>
							<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
						</Properties>
					</Attribute>
					<Attribute uuid="b1000023-0000-0000-0000-000000000001">
						<Properties>
							<Name>Статья</Name>
							<Synonym><v8:item><v8:lang>ru</v8:lang><v8:content>Статья</v8:content></v8:item></Synonym>
							<Comment/>
							<Type><v8:Type>cfg:ChartOfCharacteristicTypesRef.СтатьиАктивовПассивов</v8:Type></Type>
							<PasswordMode>false</PasswordMode><Format/><EditFormat/><ToolTip/>
							<MarkNegatives>false</MarkNegatives><Mask/><MultiLine>false</MultiLine>
							<ExtendedEdit>false</ExtendedEdit><MinValue xsi:nil="true"/><MaxValue xsi:nil="true"/>
							<FillFromFillingValue>false</FillFromFillingValue><FillValue xsi:nil="true"/>
							<FillChecking>DontCheck</FillChecking><ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
							<ChoiceParameterLinks/><ChoiceParameters/><QuickChoice>Auto</QuickChoice>
							<CreateOnInput>Auto</CreateOnInput><ChoiceForm/><LinkByType/>
							<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
						</Properties>
					</Attribute>
					<Attribute uuid="b1000024-0000-0000-0000-000000000001">
						<Properties>
							<Name>Source</Name>
							<Synonym><v8:item><v8:lang>ru</v8:lang><v8:content>Source</v8:content></v8:item></Synonym>
							<Comment/>
							<Type><v8:Type>cfg:EnumRef.ИсточникиУправленческогоБаланса</v8:Type></Type>
							<PasswordMode>false</PasswordMode><Format/><EditFormat/><ToolTip/>
							<MarkNegatives>false</MarkNegatives><Mask/><MultiLine>false</MultiLine>
							<ExtendedEdit>false</ExtendedEdit><MinValue xsi:nil="true"/><MaxValue xsi:nil="true"/>
							<FillFromFillingValue>false</FillFromFillingValue><FillValue xsi:nil="true"/>
							<FillChecking>DontCheck</FillChecking><ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
							<ChoiceParameterLinks/><ChoiceParameters/><QuickChoice>Auto</QuickChoice>
							<CreateOnInput>Auto</CreateOnInput><ChoiceForm/><LinkByType/>
							<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
						</Properties>
					</Attribute>
					<Attribute uuid="b1000025-0000-0000-0000-000000000001">
						<Properties>
							<Name>ПАП_Sign</Name>
							<Synonym><v8:item><v8:lang>ru</v8:lang><v8:content>ПАП (signed)</v8:content></v8:item></Synonym>
							<Comment/>
							<Type><v8:Type>xs:decimal</v8:Type><v8:NumberQualifiers><v8:Digits>15</v8:Digits><v8:FractionDigits>2</v8:FractionDigits><v8:AllowedSign>Any</v8:AllowedSign></v8:NumberQualifiers></Type>
							<PasswordMode>false</PasswordMode><Format/><EditFormat/><ToolTip/>
							<MarkNegatives>false</MarkNegatives><Mask/><MultiLine>false</MultiLine>
							<ExtendedEdit>false</ExtendedEdit><MinValue xsi:nil="true"/><MaxValue xsi:nil="true"/>
							<FillFromFillingValue>false</FillFromFillingValue><FillValue xsi:nil="true"/>
							<FillChecking>DontCheck</FillChecking><ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
							<ChoiceParameterLinks/><ChoiceParameters/><QuickChoice>Auto</QuickChoice>
							<CreateOnInput>Auto</CreateOnInput><ChoiceForm/><LinkByType/>
							<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
						</Properties>
					</Attribute>
					<Attribute uuid="b1000026-0000-0000-0000-000000000001">
						<Properties>
							<Name>РСППС_Sign</Name>
							<Synonym><v8:item><v8:lang>ru</v8:lang><v8:content>РСППС (inv signed)</v8:content></v8:item></Synonym>
							<Comment/>
							<Type><v8:Type>xs:decimal</v8:Type><v8:NumberQualifiers><v8:Digits>15</v8:Digits><v8:FractionDigits>2</v8:FractionDigits><v8:AllowedSign>Any</v8:AllowedSign></v8:NumberQualifiers></Type>
							<PasswordMode>false</PasswordMode><Format/><EditFormat/><ToolTip/>
							<MarkNegatives>false</MarkNegatives><Mask/><MultiLine>false</MultiLine>
							<ExtendedEdit>false</ExtendedEdit><MinValue xsi:nil="true"/><MaxValue xsi:nil="true"/>
							<FillFromFillingValue>false</FillFromFillingValue><FillValue xsi:nil="true"/>
							<FillChecking>DontCheck</FillChecking><ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
							<ChoiceParameterLinks/><ChoiceParameters/><QuickChoice>Auto</QuickChoice>
							<CreateOnInput>Auto</CreateOnInput><ChoiceForm/><LinkByType/>
							<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
						</Properties>
					</Attribute>
					<Attribute uuid="b1000027-0000-0000-0000-000000000001">
						<Properties>
							<Name>Дельта</Name>
							<Synonym><v8:item><v8:lang>ru</v8:lang><v8:content>Дельта</v8:content></v8:item></Synonym>
							<Comment/>
							<Type><v8:Type>xs:decimal</v8:Type><v8:NumberQualifiers><v8:Digits>15</v8:Digits><v8:FractionDigits>2</v8:FractionDigits><v8:AllowedSign>Any</v8:AllowedSign></v8:NumberQualifiers></Type>
							<PasswordMode>false</PasswordMode><Format/><EditFormat/><ToolTip/>
							<MarkNegatives>false</MarkNegatives><Mask/><MultiLine>false</MultiLine>
							<ExtendedEdit>false</ExtendedEdit><MinValue xsi:nil="true"/><MaxValue xsi:nil="true"/>
							<FillFromFillingValue>false</FillFromFillingValue><FillValue xsi:nil="true"/>
							<FillChecking>DontCheck</FillChecking><ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
							<ChoiceParameterLinks/><ChoiceParameters/><QuickChoice>Auto</QuickChoice>
							<CreateOnInput>Auto</CreateOnInput><ChoiceForm/><LinkByType/>
							<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
						</Properties>
					</Attribute>
					<Attribute uuid="b1000028-0000-0000-0000-000000000001">
						<Properties>
							<Name>НовоеСостояние</Name>
							<Synonym><v8:item><v8:lang>ru</v8:lang><v8:content>Новое состояние</v8:content></v8:item></Synonym>
							<Comment/>
							<Type><v8:Type>xs:string</v8:Type><v8:StringQualifiers><v8:Length>100</v8:Length><v8:AllowedLength>Variable</v8:AllowedLength></v8:StringQualifiers></Type>
							<PasswordMode>false</PasswordMode><Format/><EditFormat/><ToolTip/>
							<MarkNegatives>false</MarkNegatives><Mask/><MultiLine>false</MultiLine>
							<ExtendedEdit>false</ExtendedEdit><MinValue xsi:nil="true"/><MaxValue xsi:nil="true"/>
							<FillFromFillingValue>false</FillFromFillingValue><FillValue xsi:nil="true"/>
							<FillChecking>DontCheck</FillChecking><ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
							<ChoiceParameterLinks/><ChoiceParameters/><QuickChoice>Auto</QuickChoice>
							<CreateOnInput>Auto</CreateOnInput><ChoiceForm/><LinkByType/>
							<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
						</Properties>
					</Attribute>
					<Attribute uuid="b1000029-0000-0000-0000-000000000001">
						<Properties>
							<Name>ОшибкаТекст</Name>
							<Synonym><v8:item><v8:lang>ru</v8:lang><v8:content>Текст ошибки</v8:content></v8:item></Synonym>
							<Comment/>
							<Type><v8:Type>xs:string</v8:Type><v8:StringQualifiers><v8:Length>500</v8:Length><v8:AllowedLength>Variable</v8:AllowedLength></v8:StringQualifiers></Type>
							<PasswordMode>false</PasswordMode><Format/><EditFormat/><ToolTip/>
							<MarkNegatives>false</MarkNegatives><Mask/><MultiLine>false</MultiLine>
							<ExtendedEdit>false</ExtendedEdit><MinValue xsi:nil="true"/><MaxValue xsi:nil="true"/>
							<FillFromFillingValue>false</FillFromFillingValue><FillValue xsi:nil="true"/>
							<FillChecking>DontCheck</FillChecking><ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
							<ChoiceParameterLinks/><ChoiceParameters/><QuickChoice>Auto</QuickChoice>
							<CreateOnInput>Auto</CreateOnInput><ChoiceForm/><LinkByType/>
							<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
						</Properties>
					</Attribute>
				</ChildObjects>
			</TabularSection>
			<Form>Форма</Form>
		</ChildObjects>
	</ExternalReport>
</MetaDataObject>
```

- [ ] **Step 4: Запустить /erf-validate чтобы убедиться что XML валиден**

Run: invoke skill `/erf-validate` с путём `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса`

Expected: «✓ Валидация прошла» (или эквивалент). Если ошибки UUID-дубль — заменить UUID в новых блоках на random.

- [ ] **Step 5: Commit XML обновление**

```bash
git add _Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса.xml
git commit -m "$(cat <<'EOF'
feat(obrabotka_disbalansa_statyam): XML +2 реквизита +ТЧ ДокументыРасхождения

- Добавил реквизиты: Подразделение, ПоказыватьВсе
- Добавил ТЧ ДокументыРасхождения (10 колонок: Обработан/Документ/
  Подразделение/Статья/Source/ПАП_Sign/РСППС_Sign/Дельта/НовоеСостояние/
  ОшибкаТекст)

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 3: Pre-test SQL запросов через MCP

**Files:**
- None (только MCP execute_query)

Перед написанием BSL — прогоняем оба ключевых SQL запроса через MCP `execute_query` (memory `feedback_query_pre_check` + Rule #-1). Если падают — фиксим тут.

- [ ] **Step 1: Pre-test запрос плугов из А_ОтчетБаланс_Свод**

Use tool `mcp__1c-workerp__execute_query`:

```sql
ВЫБРАТЬ РАЗЛИЧНЫЕ
    Т.Регистратор.Месяц КАК Месяц,
    Т.Подразделение.Код КАК ПодрКод,
    Т.Подразделение КАК Подр,
    Т.Статья КАК Статья,
    Т.Source КАК Source
ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Т
ГДЕ Т.Расхождение = ИСТИНА
    И Т.Регистратор.Месяц МЕЖДУ ДАТАВРЕМЯ(2026,4,1) И ДАТАВРЕМЯ(2026,4,30,23,59,59)
    И Т.Source В (
        ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСКлиентамиПоСрокам),
        ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСПоставщикамиПоСрокам))
    И Т.Организация.КодПоЕДРПОУ = "40645273"
```

Expected: не падает, возвращает ≥ 9 строк (наши 9 подразделений ЗПП за апрель + возможно ещё ЗадКл с РСКПС).

- [ ] **Step 2: Pre-test финального запроса (объединённый ПАП vs РС с разделением РСКПС/РСППС по знакам)**

Use tool `mcp__1c-workerp__execute_query`:

```sql
ВЫБРАТЬ
    Т.Регистратор КАК Док,
    ПРЕДСТАВЛЕНИЕ(Т.Регистратор) КАК ДокИмя,
    Т.Подразделение.Код КАК ПодрКод,
    Т.Подразделение КАК Подр,
    Т.Статья КАК Статья,
    Т.Источник КАК Source,
    СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                ТОГДА Т.Сумма ИНАЧЕ -Т.Сумма КОНЕЦ) КАК ПАП_Sign
ПОМЕСТИТЬ втПАП
ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
ГДЕ Т.Период МЕЖДУ ДАТАВРЕМЯ(2026,4,1) И ДАТАВРЕМЯ(2026,4,30,23,59,59)
    И Т.Организация.КодПоЕДРПОУ = "40645273"
    И Т.Подразделение.Код В ("00-001029","00-001022","00-001026","0Ц-000004","00-001003","0Ц-000001","00-000210","00-000219","00-000190")
    И Т.Источник = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСПоставщикамиПоСрокам)
    И Т.Статья = ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьПередПоставщиками)
СГРУППИРОВАТЬ ПО Т.Регистратор, ПРЕДСТАВЛЕНИЕ(Т.Регистратор), Т.Подразделение.Код, Т.Подразделение, Т.Статья, Т.Источник
;
ВЫБРАТЬ
    Р.ДокументРегистратор КАК Док,
    ПРЕДСТАВЛЕНИЕ(Р.ДокументРегистратор) КАК ДокИмя,
    ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
        ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
        ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
    КОНЕЦ КАК ПодрКод,
    ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьПередПоставщиками) КАК Статья,
    ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСПоставщикамиПоСрокам) КАК Source,
    -СУММА(ВЫБОР КОГДА Р.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
                ТОГДА Р.ДолгУпр ИНАЧЕ -Р.ДолгУпр КОНЕЦ) КАК РСППС_Inv
ПОМЕСТИТЬ втРС
ИЗ РегистрНакопления.РасчетыСПоставщикамиПоСрокам КАК Р
    ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
        ПО Р.АналитикаУчетаПоПартнерам = АП.КлючАналитики
ГДЕ Р.Период МЕЖДУ ДАТАВРЕМЯ(2026,4,1) И ДАТАВРЕМЯ(2026,4,30,23,59,59)
    И АП.Организация.КодПоЕДРПОУ = "40645273"
    И АП.Партнер <> ЗНАЧЕНИЕ(Справочник.Партнеры.НашеПредприятие)
    И ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
        ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
        ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
    КОНЕЦ В ("00-001029","00-001022","00-001026","0Ц-000004","00-001003","0Ц-000001","00-000210","00-000219","00-000190")
СГРУППИРОВАТЬ ПО
    Р.ДокументРегистратор, ПРЕДСТАВЛЕНИЕ(Р.ДокументРегистратор),
    ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
        ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
        ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
    КОНЕЦ
;
ВЫБРАТЬ
    ЕСТЬNULL(ПАП.ПодрКод, РС.ПодрКод) КАК ПодрКод,
    ЕСТЬNULL(ПАП.ДокИмя, РС.ДокИмя) КАК ДокИмя,
    ЕСТЬNULL(ПАП.ПАП_Sign, 0) КАК ПАП_Sign,
    ЕСТЬNULL(РС.РСППС_Inv, 0) КАК РСППС_Sign,
    ЕСТЬNULL(ПАП.ПАП_Sign, 0) - ЕСТЬNULL(РС.РСППС_Inv, 0) КАК Дельта
ИЗ втПАП КАК ПАП
    ПОЛНОЕ СОЕДИНЕНИЕ втРС КАК РС
    ПО ПАП.Док = РС.Док И ПАП.ПодрКод = РС.ПодрКод
ГДЕ ЕСТЬNULL(ПАП.ПАП_Sign, 0) - ЕСТЬNULL(РС.РСППС_Inv, 0) > 0.01
   ИЛИ ЕСТЬNULL(ПАП.ПАП_Sign, 0) - ЕСТЬNULL(РС.РСППС_Inv, 0) < -0.01
УПОРЯДОЧИТЬ ПО ПодрКод, Дельта
```

Expected: 19 строк, Σ Δ по 9 подр совпадает с baseline (из Task 1).

**Если результат отличается от baseline (например, пользователь успел перепровести часть)** — это нормально, базовый запрос работает; примем live snapshot. Главное чтобы синтаксис не падал.

- [ ] **Step 3: Зафиксировать что pre-test пройден**

Никаких файлов. Просто продолжаем — оба запроса валидны, можно переносить в BSL.

---

### Task 4: ObjectModule.bsl — структура и helper-функции

**Files:**
- Create: `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Ext/ObjectModule.bsl`

Файл сейчас почти пустой (1 строка). Создаём с нуля. Декомпозиция:
- `АнализРасхождений()` Экспорт — точка входа кнопки 1
- `ПерепровестиДокументы()` Экспорт — точка входа кнопки 2
- `_ПолучитьОрганизациюПоУмолчанию()` — helper
- `_ЗапросПлугов(НачПериода, КонПериода, ВыбПодр)` — возвращает ТЗ месяцев с плугами
- `_ЗапросВиновниковМесяца(НачМес, КонМес, МасивКодов)` — FULL OUTER ПАП vs РС
- `_ПересчитатьДельту(Док, Подр, Статья, Source)` — мини-запрос для одной строки

- [ ] **Step 1: Создать ObjectModule.bsl с шапкой и helpers**

Create file `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Ext/ObjectModule.bsl`:

```bsl
// ============================================================================
// А_ОбработкаДисбалансаПоСтатьямБаланса — Модуль объекта
// ============================================================================
// Поиск и устранение Расхождение=Истина в РегистрСведений.А_ОтчетБаланс_Свод
// по Source ∈ {РасчетыСКлиентамиПоСрокам, РасчетыСПоставщикамиПоСрокам}.
//
// Кнопка "Анализ расхождений" (АнализРасхождений):
//   1. Запрос плугов из А_ОтчетБаланс_Свод за период (Расхождение=Истина,
//      Source РСКПС/РСППС, фильтр по орг + опц. подр)
//   2. Per месяц — FULL OUTER ПАП vs РСКПС/РСППС через ДокументРегистратор
//   3. Документы с |Δ|>0.01 → ТЧ ДокументыРасхождения
//
// Кнопка "Перепровести" (ПерепровестиДокументы):
//   Для каждой строки ТЧ с Обработан=Ложь — транзакция per-документ:
//   НачатьТранзакцию → Записать(Проведение) → ЗафиксироватьТранзакцию
//   При ошибке — ОтменитьТранзакцию + ОшибкаТекст в ТЧ.
//   После — пересчёт Δ через _ПересчитатьДельту → Обработан=Истина если Δ=0.
//
// Spec: docs/superpowers/specs/2026-05-25-obrabotka-disbalansa-statyam-design.md
// ============================================================================

Функция _ПолучитьОрганизациюПоУмолчанию()
	// ТОВ ІНДАСТРІАЛБУД (ЕДРПОУ 40645273) — единственная нужная орг для balans
	Возврат Справочники.Организации.НайтиПоРеквизиту("КодПоЕДРПОУ", "40645273");
КонецФункции

Функция _ИсточникРСКПС() Экспорт
	Возврат Перечисления.ИсточникиУправленческогоБаланса.РасчетыСКлиентамиПоСрокам;
КонецФункции

Функция _ИсточникРСППС() Экспорт
	Возврат Перечисления.ИсточникиУправленческогоБаланса.РасчетыСПоставщикамиПоСрокам;
КонецФункции

Функция _СтатьиПоСорсу(Источник)
	// Возвращает список статей для заданного Source
	// (по канону Свод_РасчетыСПартнерами в А_ФинРез_Баланс.ObjectModule.bsl)
	Сп = Новый СписокЗначений;
	Если Источник = _ИсточникРСППС() Тогда
		Сп.Добавить(ПланыВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьПередПоставщиками);
		Сп.Добавить(ПланыВидовХарактеристик.СтатьиАктивовПассивов.ВыданныеАвансы);
		Сп.Добавить(ПланыВидовХарактеристик.СтатьиАктивовПассивов.ОбязательстваПередСобственнымиОрганизациями);
		Сп.Добавить(ПланыВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьСобственныхОрганизаций);
	Иначе  // РСКПС
		Сп.Добавить(ПланыВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьКлиентов);
		Сп.Добавить(ПланыВидовХарактеристик.СтатьиАктивовПассивов.ПолученныеАвансы);
		Сп.Добавить(ПланыВидовХарактеристик.СтатьиАктивовПассивов.ОбязательстваПередСобственнымиОрганизациями);
		Сп.Добавить(ПланыВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьСобственныхОрганизаций);
	КонецЕсли;
	Возврат Сп;
КонецФункции
```

- [ ] **Step 2: Добавить функцию _ЗапросПлугов (поиск плугов из А_ОтчетБаланс_Свод)**

Append to `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Ext/ObjectModule.bsl`:

```bsl

Функция _ЗапросПлугов(Орг, НачПериода, КонПериода, ВыбПодр)
	// Возвращает ТЗ: Месяц / ПодрКод / Подр / Статья / Source
	// для всех плугов Расхождение=Истина по Source ∈ {РСКПС, РСППС}
	Запрос = Новый Запрос;
	Запрос.УстановитьПараметр("Орг", Орг);
	Запрос.УстановитьПараметр("НачПериода", НачалоДня(НачПериода));
	Запрос.УстановитьПараметр("КонПериода", КонецДня(КонПериода));
	Запрос.УстановитьПараметр("РСКПС", _ИсточникРСКПС());
	Запрос.УстановитьПараметр("РСППС", _ИсточникРСППС());
	Запрос.УстановитьПараметр("ВыбПодр", ВыбПодр);
	Запрос.УстановитьПараметр("ИспользоватьФильтрПодр", ЗначениеЗаполнено(ВыбПодр));
	Запрос.Текст =
	"ВЫБРАТЬ РАЗЛИЧНЫЕ
	|	Т.Регистратор.Месяц КАК Месяц,
	|	Т.Подразделение.Код КАК ПодрКод,
	|	Т.Подразделение КАК Подр,
	|	Т.Статья КАК Статья,
	|	Т.Source КАК Source
	|ИЗ
	|	РегистрСведений.А_ОтчетБаланс_Свод КАК Т
	|ГДЕ
	|	Т.Расхождение = ИСТИНА
	|	И Т.Регистратор.Месяц МЕЖДУ &НачПериода И &КонПериода
	|	И Т.Source В (&РСКПС, &РСППС)
	|	И Т.Организация = &Орг
	|	И (НЕ &ИспользоватьФильтрПодр ИЛИ Т.Подразделение В ИЕРАРХИИ (&ВыбПодр))
	|УПОРЯДОЧИТЬ ПО Месяц, ПодрКод, Статья";
	Возврат Запрос.Выполнить().Выгрузить();
КонецФункции

Функция _ПервыйДеньМесяца(Дата)
	Возврат НачалоМесяца(Дата);
КонецФункции

Функция _ПоследнийДеньМесяца(Дата)
	Возврат КонецМесяца(Дата);
КонецФункции
```

- [ ] **Step 3: Commit базового каркаса**

```bash
git add _Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Ext/ObjectModule.bsl
git commit -m "$(cat <<'EOF'
feat(obrabotka_disbalansa_statyam): ObjectModule.bsl — helper'ы

Каркас ObjectModule.bsl с helper-функциями:
- _ПолучитьОрганизациюПоУмолчанию (ТОВ ІНДАСТРІАЛБУД по ЕДРПОУ)
- _ИсточникРСКПС/_ИсточникРСППС (константы перечисления)
- _СтатьиПоСорсу (список статей для Source согласно canon)
- _ЗапросПлугов (поиск плугов из А_ОтчетБаланс_Свод)

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 5: ObjectModule.bsl — функция _ЗапросВиновниковМесяца (FULL OUTER)

**Files:**
- Modify: `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Ext/ObjectModule.bsl`

Это сердце анализа — FULL OUTER ПАП vs РСКПС/РСППС за один месяц. Делаем универсальную функцию которая работает и для РСКПС и для РСППС (через UNION ALL и инверсию знака для поставщиков).

- [ ] **Step 1: Добавить функцию _ЗапросВиновниковМесяца**

Append to `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Ext/ObjectModule.bsl`:

```bsl

Функция _ЗапросВиновниковМесяца(Орг, НачМес, КонМес, МасивКодов, МасивСтатей)
	// FULL OUTER JOIN ПАП vs (РСКПС + РСППС) для одного месяца.
	//
	// Знак Δ зафиксирован практикой апреля 2026 (см. spec §4.3):
	//   - Для РСППС (поставщики): инвертируем -ДолгУпр в стороне РС
	//   - Для РСКПС (клиенты): не инвертируем
	//   Тогда Δ = ПАП_Sign - РС_Sign (где РС_Sign уже содержит правильный знак)
	//
	// Возвращает ТЗ: Док / ДокИмя / ПодрКод / Подр / Статья / Source /
	//                ПАП_Sign / РС_Sign / Дельта
	Запрос = Новый Запрос;
	Запрос.УстановитьПараметр("Орг", Орг);
	Запрос.УстановитьПараметр("НачМес", НачалоДня(НачМес));
	Запрос.УстановитьПараметр("КонМес", КонецДня(КонМес));
	Запрос.УстановитьПараметр("МасивКодов", МасивКодов);
	Запрос.УстановитьПараметр("МасивСтатей", МасивСтатей);
	Запрос.УстановитьПараметр("РСКПС", _ИсточникРСКПС());
	Запрос.УстановитьПараметр("РСППС", _ИсточникРСППС());
	Запрос.УстановитьПараметр("НашеПредприятие", Справочники.Партнеры.НашеПредприятие);
	Запрос.Текст =
	"ВЫБРАТЬ
	|	Т.Регистратор КАК Док,
	|	ПРЕДСТАВЛЕНИЕ(Т.Регистратор) КАК ДокИмя,
	|	Т.Подразделение.Код КАК ПодрКод,
	|	Т.Подразделение КАК Подр,
	|	Т.Статья КАК Статья,
	|	Т.Источник КАК Source,
	|	СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
	|				ТОГДА Т.Сумма ИНАЧЕ -Т.Сумма КОНЕЦ) КАК ПАП_Sign
	|ПОМЕСТИТЬ втПАП
	|ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
	|ГДЕ Т.Период МЕЖДУ &НачМес И &КонМес
	|	И Т.Организация = &Орг
	|	И Т.Подразделение.Код В (&МасивКодов)
	|	И Т.Источник В (&РСКПС, &РСППС)
	|	И Т.Статья В (&МасивСтатей)
	|СГРУППИРОВАТЬ ПО Т.Регистратор, Т.Подразделение.Код, Т.Подразделение, Т.Статья, Т.Источник
	|;
	|// РСКПС-ветка (клиенты) — БЕЗ инверсии ДолгУпр для статей долга клиентов
	|ВЫБРАТЬ
	|	Р.ДокументРегистратор КАК Док,
	|	ПРЕДСТАВЛЕНИЕ(Р.ДокументРегистратор) КАК ДокИмя,
	|	ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
	|		ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
	|		ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
	|	КОНЕЦ КАК ПодрКод,
	|	ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
	|		ТОГДА Р.ОбъектРасчетов.Договор.Подразделение
	|		ИНАЧЕ Р.ОбъектРасчетов.Подразделение
	|	КОНЕЦ КАК Подр,
	|	&РСКПС КАК Source,
	|	СУММА(ВЫБОР КОГДА Р.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
	|				ТОГДА Р.ДолгУпр ИНАЧЕ -Р.ДолгУпр КОНЕЦ) КАК РС_Sign
	|ПОМЕСТИТЬ втРС
	|ИЗ РегистрНакопления.РасчетыСКлиентамиПоСрокам КАК Р
	|	ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
	|		ПО Р.АналитикаУчетаПоПартнерам = АП.КлючАналитики
	|ГДЕ Р.Период МЕЖДУ &НачМес И &КонМес
	|	И АП.Организация = &Орг
	|	И АП.Партнер <> &НашеПредприятие
	|	И ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
	|		ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
	|		ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
	|	КОНЕЦ В (&МасивКодов)
	|СГРУППИРОВАТЬ ПО
	|	Р.ДокументРегистратор,
	|	ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
	|		ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
	|		ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
	|	КОНЕЦ,
	|	ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
	|		ТОГДА Р.ОбъектРасчетов.Договор.Подразделение
	|		ИНАЧЕ Р.ОбъектРасчетов.Подразделение
	|	КОНЕЦ
	|
	|ОБЪЕДИНИТЬ ВСЕ
	|
	|// РСППС-ветка (поставщики) — С ИНВЕРСИЕЙ ДолгУпр
	|ВЫБРАТЬ
	|	Р.ДокументРегистратор,
	|	ПРЕДСТАВЛЕНИЕ(Р.ДокументРегистратор),
	|	ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
	|		ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
	|		ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
	|	КОНЕЦ,
	|	ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
	|		ТОГДА Р.ОбъектРасчетов.Договор.Подразделение
	|		ИНАЧЕ Р.ОбъектРасчетов.Подразделение
	|	КОНЕЦ,
	|	&РСППС,
	|	-СУММА(ВЫБОР КОГДА Р.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
	|				ТОГДА Р.ДолгУпр ИНАЧЕ -Р.ДолгУпр КОНЕЦ)
	|ИЗ РегистрНакопления.РасчетыСПоставщикамиПоСрокам КАК Р
	|	ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
	|		ПО Р.АналитикаУчетаПоПартнерам = АП.КлючАналитики
	|ГДЕ Р.Период МЕЖДУ &НачМес И &КонМес
	|	И АП.Организация = &Орг
	|	И АП.Партнер <> &НашеПредприятие
	|	И ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
	|		ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
	|		ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
	|	КОНЕЦ В (&МасивКодов)
	|СГРУППИРОВАТЬ ПО
	|	Р.ДокументРегистратор,
	|	ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
	|		ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
	|		ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
	|	КОНЕЦ,
	|	ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
	|		ТОГДА Р.ОбъектРасчетов.Договор.Подразделение
	|		ИНАЧЕ Р.ОбъектРасчетов.Подразделение
	|	КОНЕЦ
	|;
	|// FULL OUTER + фильтр |Δ|>0.01
	|ВЫБРАТЬ
	|	ЕСТЬNULL(ПАП.Док, РС.Док) КАК Док,
	|	ЕСТЬNULL(ПАП.ДокИмя, РС.ДокИмя) КАК ДокИмя,
	|	ЕСТЬNULL(ПАП.ПодрКод, РС.ПодрКод) КАК ПодрКод,
	|	ЕСТЬNULL(ПАП.Подр, РС.Подр) КАК Подр,
	|	ЕСТЬNULL(ПАП.Статья, РС.Статья) КАК Статья,
	|	ЕСТЬNULL(ПАП.Source, РС.Source) КАК Source,
	|	ЕСТЬNULL(ПАП.ПАП_Sign, 0) КАК ПАП_Sign,
	|	ЕСТЬNULL(РС.РС_Sign, 0) КАК РС_Sign,
	|	ЕСТЬNULL(ПАП.ПАП_Sign, 0) - ЕСТЬNULL(РС.РС_Sign, 0) КАК Дельта
	|ИЗ втПАП КАК ПАП
	|	ПОЛНОЕ СОЕДИНЕНИЕ втРС КАК РС
	|	ПО ПАП.Док = РС.Док И ПАП.ПодрКод = РС.ПодрКод И ПАП.Source = РС.Source
	|ГДЕ ЕСТЬNULL(ПАП.ПАП_Sign, 0) - ЕСТЬNULL(РС.РС_Sign, 0) > 0.01
	|   ИЛИ ЕСТЬNULL(ПАП.ПАП_Sign, 0) - ЕСТЬNULL(РС.РС_Sign, 0) < -0.01
	|УПОРЯДОЧИТЬ ПО ПодрКод, Дельта";
	Возврат Запрос.Выполнить().Выгрузить();
КонецФункции
```

- [ ] **Step 2: Commit FULL OUTER функцию**

```bash
git add _Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Ext/ObjectModule.bsl
git commit -m "$(cat <<'EOF'
feat(obrabotka_disbalansa_statyam): _ЗапросВиновниковМесяца (FULL OUTER ПАП vs РСКПС/РСППС)

Универсальный запрос находит первичные документы с расхождением между ПАП
и РСКПС/РСППС через ДокументРегистратор. РСППС с инверсией ДолгУпр
(canon Свод_РасчетыСПартнерами), РСКПС без инверсии.

Δ = ПАП_Sign - РС_Sign (где РС_Sign содержит правильный знак для статьи).
Фильтр |Δ|>0.01.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 6: ObjectModule.bsl — АнализРасхождений() и _ПересчитатьДельту()

**Files:**
- Modify: `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Ext/ObjectModule.bsl`

- [ ] **Step 1: Добавить процедуру АнализРасхождений**

Append:

```bsl

Процедура АнализРасхождений() Экспорт
	ДокументыРасхождения.Очистить();
	
	Если НЕ ЗначениеЗаполнено(НачалоПериода) ИЛИ НЕ ЗначениеЗаполнено(ОкончаниеПериода) Тогда
		Сообщить("Не заполнен период!");
		Возврат;
	КонецЕсли;
	
	Орг = _ПолучитьОрганизациюПоУмолчанию();
	Если НЕ ЗначениеЗаполнено(Орг) Тогда
		Сообщить("Не найдена организация по ЕДРПОУ 40645273");
		Возврат;
	КонецЕсли;
	
	// Шаг 1: Найти все плуги (Расхождение=Истина) по РСКПС/РСППС
	ТЗПлугов = _ЗапросПлугов(Орг, НачалоПериода, ОкончаниеПериода, Подразделение);
	Сообщить("Найдено ключей плугов: " + ТЗПлугов.Количество());
	
	Если ТЗПлугов.Количество() = 0 Тогда
		Сообщить("Плугов Расхождение=Истина за период не найдено — анализ не требуется.");
		Возврат;
	КонецЕсли;
	
	// Шаг 2: Сгруппировать плуги по Месяцу — для каждого месяца один запрос виновников
	МесяцыИмКлючи = Новый Соответствие; // Дата(нач месяца) → Структура{КонМес, Коды, Статьи}
	Для Каждого СтрПлуга Из ТЗПлугов Цикл
		МесКлюч = НачалоМесяца(СтрПлуга.Месяц);
		Если МесяцыИмКлючи.Получить(МесКлюч) = Неопределено Тогда
			МесяцыИмКлючи.Вставить(МесКлюч, Новый Структура(
				"КонМес, Коды, Статьи",
				КонецМесяца(СтрПлуга.Месяц),
				Новый Массив, Новый Массив));
		КонецЕсли;
		Данные = МесяцыИмКлючи.Получить(МесКлюч);
		Если Данные.Коды.Найти(СтрПлуга.ПодрКод) = Неопределено Тогда
			Данные.Коды.Добавить(СтрПлуга.ПодрКод);
		КонецЕсли;
		Если Данные.Статьи.Найти(СтрПлуга.Статья) = Неопределено Тогда
			Данные.Статьи.Добавить(СтрПлуга.Статья);
		КонецЕсли;
	КонецЦикла;
	
	// Шаг 3: По каждому месяцу — FULL OUTER, собрать виновников в ТЧ
	ВсегоВиновников = 0;
	Для Каждого КЗ Из МесяцыИмКлючи Цикл
		НачМес = КЗ.Ключ;
		КонМес = КЗ.Значение.КонМес;
		Коды = КЗ.Значение.Коды;
		Статьи = КЗ.Значение.Статьи;
		
		ТЗВинов = _ЗапросВиновниковМесяца(Орг, НачМес, КонМес, Коды, Статьи);
		
		Для Каждого СтрВ Из ТЗВинов Цикл
			НС = ДокументыРасхождения.Добавить();
			НС.Обработан = Ложь;
			НС.Документ = СтрВ.Док;
			НС.Подразделение = СтрВ.Подр;
			НС.Статья = СтрВ.Статья;
			НС.Source = СтрВ.Source;
			НС.ПАП_Sign = СтрВ.ПАП_Sign;
			НС.РСППС_Sign = СтрВ.РС_Sign;
			НС.Дельта = СтрВ.Дельта;
			НС.НовоеСостояние = "";
			НС.ОшибкаТекст = "";
			ВсегоВиновников = ВсегоВиновников + 1;
		КонецЦикла;
	КонецЦикла;
	
	// Шаг 4: Подвести итог
	СуммаΔAbs = 0;
	Для Каждого Стр Из ДокументыРасхождения Цикл
		СуммаΔAbs = СуммаΔAbs + ?(Стр.Дельта < 0, -Стр.Дельта, Стр.Дельта);
	КонецЦикла;
	Сообщить("Анализ завершён: " + ВсегоВиновников + " документов, |Σ Δ| = "
		+ Формат(СуммаΔAbs, "ЧДЦ=2"));
КонецПроцедуры
```

- [ ] **Step 2: Добавить функцию _ПересчитатьДельту**

Append:

```bsl

Функция _ПересчитатьДельту(Док, Подр, Источник) Экспорт
	// Пересчитывает Δ для одного документа после перепроведения.
	// Возвращает Структуру {ПАП, РС, Дельта}.
	Орг = _ПолучитьОрганизациюПоУмолчанию();
	НачМес = НачалоМесяца(Док.Дата);
	КонМес = КонецМесяца(Док.Дата);
	
	Запрос = Новый Запрос;
	Запрос.УстановитьПараметр("Док", Док);
	Запрос.УстановитьПараметр("Орг", Орг);
	Запрос.УстановитьПараметр("НачМес", НачМес);
	Запрос.УстановитьПараметр("КонМес", КонМес);
	Запрос.УстановитьПараметр("ПодрКод", Подр.Код);
	Запрос.УстановитьПараметр("Ист", Источник);
	Запрос.УстановитьПараметр("РСППС", _ИсточникРСППС());
	Запрос.УстановитьПараметр("Статьи", _СтатьиПоСорсу(Источник));
	Запрос.УстановитьПараметр("НашеПредприятие", Справочники.Партнеры.НашеПредприятие);
	
	// ПАП-сторона
	Запрос.Текст =
	"ВЫБРАТЬ
	|	ЕСТЬNULL(СУММА(ВЫБОР КОГДА Т.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
	|				ТОГДА Т.Сумма ИНАЧЕ -Т.Сумма КОНЕЦ), 0) КАК Sig
	|ИЗ РегистрНакопления.ПрочиеАктивыПассивы КАК Т
	|ГДЕ Т.Регистратор = &Док
	|	И Т.Подразделение.Код = &ПодрКод
	|	И Т.Источник = &Ист
	|	И Т.Статья В (&Статьи)";
	Выб = Запрос.Выполнить().Выбрать();
	Выб.Следующий();
	ПАП_S = Выб.Sig;
	
	// РС-сторона — выбираем регистр по Source
	Если Источник = _ИсточникРСППС() Тогда
		Запрос.Текст =
		"ВЫБРАТЬ
		|	-ЕСТЬNULL(СУММА(ВЫБОР КОГДА Р.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
		|					ТОГДА Р.ДолгУпр ИНАЧЕ -Р.ДолгУпр КОНЕЦ), 0) КАК Sig
		|ИЗ РегистрНакопления.РасчетыСПоставщикамиПоСрокам КАК Р
		|	ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
		|		ПО Р.АналитикаУчетаПоПартнерам = АП.КлючАналитики
		|ГДЕ Р.ДокументРегистратор = &Док
		|	И АП.Организация = &Орг
		|	И АП.Партнер <> &НашеПредприятие
		|	И ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
		|		ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
		|		ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
		|	КОНЕЦ = &ПодрКод";
	Иначе
		Запрос.Текст =
		"ВЫБРАТЬ
		|	ЕСТЬNULL(СУММА(ВЫБОР КОГДА Р.ВидДвижения = ЗНАЧЕНИЕ(ВидДвиженияНакопления.Приход)
		|					ТОГДА Р.ДолгУпр ИНАЧЕ -Р.ДолгУпр КОНЕЦ), 0) КАК Sig
		|ИЗ РегистрНакопления.РасчетыСКлиентамиПоСрокам КАК Р
		|	ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.АналитикаУчетаПоПартнерам КАК АП
		|		ПО Р.АналитикаУчетаПоПартнерам = АП.КлючАналитики
		|ГДЕ Р.ДокументРегистратор = &Док
		|	И АП.Организация = &Орг
		|	И АП.Партнер <> &НашеПредприятие
		|	И ВЫБОР КОГДА Р.ОбъектРасчетов.Подразделение = ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)
		|		ТОГДА Р.ОбъектРасчетов.Договор.Подразделение.Код
		|		ИНАЧЕ Р.ОбъектРасчетов.Подразделение.Код
		|	КОНЕЦ = &ПодрКод";
	КонецЕсли;
	Выб = Запрос.Выполнить().Выбрать();
	Выб.Следующий();
	РС_S = Выб.Sig;
	
	Возврат Новый Структура("ПАП, РС, Дельта", ПАП_S, РС_S, ПАП_S - РС_S);
КонецФункции
```

- [ ] **Step 3: Commit**

```bash
git add _Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Ext/ObjectModule.bsl
git commit -m "$(cat <<'EOF'
feat(obrabotka_disbalansa_statyam): АнализРасхождений() + _ПересчитатьДельту()

- АнализРасхождений: групирует плуги по месяцу, per-месяц вызывает
  _ЗапросВиновниковМесяца, заполняет ТЧ ДокументыРасхождения
- _ПересчитатьДельту: мини-запрос для одной строки после перепроведения
  (выбирает РСКПС/РСППС по параметру Источник)

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 7: ObjectModule.bsl — ПерепровестиДокументы()

**Files:**
- Modify: `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Ext/ObjectModule.bsl`

- [ ] **Step 1: Добавить процедуру ПерепровестиДокументы**

Append:

```bsl

Процедура ПерепровестиДокументы() Экспорт
	КолОК = 0;
	КолОшибок = 0;
	КолОсталасьΔ = 0;
	
	Для Каждого Стр Из ДокументыРасхождения Цикл
		Если Стр.Обработан Тогда
			Продолжить;
		КонецЕсли;
		Если НЕ ЗначениеЗаполнено(Стр.Документ) Тогда
			Стр.НовоеСостояние = "Пустая ссылка";
			Продолжить;
		КонецЕсли;
		
		НачатьТранзакцию();
		Попытка
			ДокОбъект = Стр.Документ.ПолучитьОбъект();
			Если ДокОбъект = Неопределено Тогда
				ВызватьИсключение "Битая ссылка / документ удалён";
			КонецЕсли;
			
			// Записать с режимом Проведение автоматически делает
			// и распроведение, и проведение в одной атомарной операции платформы.
			// Подписки событий (включая создание РегистраторРасчётов) срабатывают.
			ДокОбъект.Записать(РежимЗаписиДокумента.Проведение);
			
			ЗафиксироватьТранзакцию();
			
			// Пересчёт Δ после перепроведения
			НоваяΔ = _ПересчитатьДельту(Стр.Документ, Стр.Подразделение, Стр.Source);
			Стр.ПАП_Sign = НоваяΔ.ПАП;
			Стр.РСППС_Sign = НоваяΔ.РС;
			Стр.Дельта = НоваяΔ.Дельта;
			
			АбсΔ = ?(НоваяΔ.Дельта < 0, -НоваяΔ.Дельта, НоваяΔ.Дельта);
			Если Окр(АбсΔ, 2) < 0.01 Тогда
				Стр.Обработан = Истина;
				Стр.НовоеСостояние = "ОК";
				КолОК = КолОК + 1;
			Иначе
				Стр.НовоеСостояние = "Δ=" + Формат(НоваяΔ.Дельта, "ЧДЦ=2");
				КолОсталасьΔ = КолОсталасьΔ + 1;
			КонецЕсли;
		Исключение
			ОтменитьТранзакцию();
			Стр.НовоеСостояние = "ОШИБКА";
			Стр.ОшибкаТекст = Лев(ОписаниеОшибки(), 500);
			КолОшибок = КолОшибок + 1;
		КонецПопытки;
	КонецЦикла;
	
	Сообщить("Перепроведение завершено:");
	Сообщить("  Стало OK: " + КолОК);
	Сообщить("  Осталась Δ: " + КолОсталасьΔ);
	Сообщить("  Ошибок: " + КолОшибок);
КонецПроцедуры
```

- [ ] **Step 2: Commit**

```bash
git add _Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Ext/ObjectModule.bsl
git commit -m "$(cat <<'EOF'
feat(obrabotka_disbalansa_statyam): ПерепровестиДокументы() — per-документ транзакция

Цикл по ТЧ ДокументыРасхождения с Обработан=Ложь:
- НачатьТранзакцию → Записать(Проведение) → ЗафиксироватьТранзакцию
- При ошибке: ОтменитьТранзакцию + ОшибкаТекст в ТЧ
- После успеха: _ПересчитатьДельту → Обработан=Истина если |Δ|<0.01

Согласовано с пользователем: транзакция per-документ
(не глобальная), при ошибке откат + текст ошибки видим в ТЧ.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 8: Форма обработки — добавить колонки ТЧ + 2 кнопки

**Files:**
- Modify: `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Forms/Форма/Ext/Form.xml`
- Modify: `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Forms/Форма/Ext/Form/Module.bsl`

- [ ] **Step 1: Прочитать Form.xml образца чтобы скопировать паттерн колонок ТЧ**

Read: `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоПодразделениям/Forms/Форма/Ext/Form.xml`

Запомнить: блок `<Items>` с `Field` для каждой колонки ТЧ + блок Commands + GroupBox с кнопками.

- [ ] **Step 2: Прочитать текущую Form.xml**

Read: `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Forms/Форма/Ext/Form.xml`

Запомнить структуру. Это короткий файл (есть только период + кнопка ВыбратьПериод).

- [ ] **Step 3: Скопировать Form.xml образца как старт + точечные текстовые замены**

```bash
cp "_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоПодразделениям/Forms/Форма/Ext/Form.xml" \
   "_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Forms/Форма/Ext/Form.xml"
```

После копирования — применить следующие текстовые замены в новом файле (можно скриптом `sed -i` под bash или ручным Edit'ом — список конкретен):

1. **Имя ТЧ:** `ДокументыДисбаланса` → `ДокументыРасхождения` (все вхождения)
2. **Команды:** `АнализДокументов` → `АнализРасхождений`, `СоздатьДокументы` → `Перепровести`
3. **Лейбл кнопки:** `Создать документы` → `Перепровести`, `Анализ документов` → `Анализ расхождений`
4. **Удалить колонки которых нет в нашей ТЧ:** `ПодразделениеПриход`, `ПодразделениеРасход`, `СуммаДокумента`, `ДокВзаимозачетЗадолженности`, `СуммаВзаимозачета`, `СтатьяДисбаланса`, `СуммаРасхождение`, `РасхождениеДоговор` (8 колонок Items + соответствующие Field-блоки)
5. **Добавить новые колонки** (по аналогии с оставшимися Field-блоками, заменив только Name/PathSegments):
   - `Статья` (тип = ChartOfCharacteristicTypesRef.СтатьиАктивовПассивов)
   - `Source` (тип = EnumRef.ИсточникиУправленческогоБаланса)
   - `ПАП_Sign` (тип = Number)
   - `РСППС_Sign` (тип = Number)
   - `Дельта` (тип = Number)
   - `НовоеСостояние` (тип = String 100)
   - `ОшибкаТекст` (тип = String 500)

**Технический совет:** реализующий агент должен сначала прочитать оба файла (новый и образец), затем применить точечные `Edit` операции. Для шага 4 и 5 — использовать Edit на каждой колонке отдельно (одна колонка = два Edit-вызова: удалить старый Field-блок + удалить ссылку из Items).

После всех правок — Validate через skill `/form-validate`:

```
/form-validate _Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Forms/Форма
```

Expected: «✓ Форма валидна»

- [ ] **Step 4: Обновить Form/Module.bsl с обработчиками новых кнопок**

Replace content of `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Forms/Форма/Ext/Form/Module.bsl`:

```bsl

&НаКлиенте
Процедура ВыбратьПериод(Команда)
	Диалог = Новый ДиалогРедактированияСтандартногоПериода;
	Если ЗначениеЗаполнено(Объект.НачалоПериода) Тогда
		Диалог.Период.ДатаНачала = Объект.НачалоПериода;
	КонецЕсли;
	Если ЗначениеЗаполнено(Объект.ОкончаниеПериода) Тогда
		Диалог.Период.ДатаОкончания = Объект.ОкончаниеПериода;
	КонецЕсли;
	ОписаниеОповещения = Новый ОписаниеОповещения("ВыбратьПериодЗавершение", ЭтотОбъект);
	Диалог.Показать(ОписаниеОповещения);
КонецПроцедуры

&НаКлиенте
Процедура ВыбратьПериодЗавершение(Период, ДопПарам) Экспорт
	Если Период = Неопределено Тогда
		Возврат;
	КонецЕсли;
	Объект.НачалоПериода = Период.ДатаНачала;
	Объект.ОкончаниеПериода = Период.ДатаОкончания;
КонецПроцедуры

&НаКлиенте
Процедура АнализРасхождений(Команда)
	АнализРасхожденийНаСервере();
КонецПроцедуры

&НаСервере
Процедура АнализРасхожденийНаСервере()
	мОбъект = РеквизитФормыВЗначение("Объект");
	мОбъект.АнализРасхождений();
	ЗначениеВРеквизитФормы(мОбъект, "Объект");
КонецПроцедуры

&НаКлиенте
Процедура Перепровести(Команда)
	ПерепровестиНаСервере();
КонецПроцедуры

&НаСервере
Процедура ПерепровестиНаСервере()
	мОбъект = РеквизитФормыВЗначение("Объект");
	мОбъект.ПерепровестиДокументы();
	ЗначениеВРеквизитФормы(мОбъект, "Объект");
КонецПроцедуры
```

- [ ] **Step 5: Commit формы**

```bash
git add _Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса/Forms/Форма/
git commit -m "$(cat <<'EOF'
feat(obrabotka_disbalansa_statyam): Form.xml +Module.bsl — UI с ТЧ и 2 кнопками

- Form.xml: добавлены поля Подразделение, ПоказыватьВсе, ТЧ ДокументыРасхождения
  с 10 колонками, кнопки АнализРасхождений и Перепровести
- Form/Module.bsl: обработчики кнопок (НаКлиенте → НаСервере → метод объекта)

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 9: Сборка ERF через /erf-build

**Files:**
- Создаётся: `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса.erf`

- [ ] **Step 1: Запустить /erf-build**

Invoke skill: `/erf-build _Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса`

Expected: создан `.erf` файл рядом с XML.

- [ ] **Step 2: Убедиться что файл создан и не пустой**

Run: `ls -la _Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса.erf`

Expected: размер ≥ 5 KB (не 0).

- [ ] **Step 3: Commit ERF (если не gitignored)**

```bash
git add _Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса.erf 2>/dev/null && \
git commit -m "build: erf для obrabotka_disbalansa_statyam" || echo "ERF gitignored, skip"
```

Если .erf в .gitignore — пропустить (это нормально, ERF пересобирается).

---

### Task 10: Test 1 — Python COM вызов АнализРасхождений + проверка против baseline

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_test1_analiz.py`

- [ ] **Step 1: Написать Тест 1**

Create file `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_test1_analiz.py`:

```python
# -*- coding: utf-8 -*-
"""
TEST 1 — Анализ через Python COM вызов внешнего отчёта.

Вызывает АнализРасхождений() и сравнивает результат ТЧ с baseline
(obrabotka_baseline.json, зафиксированным в Task 1).

Acceptance: 
  - ТЧ содержит ровно 19 строк
  - Σ Δ по 9 подразделениям совпадает с baseline до копейки
"""
import sys, io, json, os, datetime as dt
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp, ARTIFACTS_DIR

ERF_PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\А_ОбработкаДисбалансаПоСтатьямБаланса.erf"
BASELINE_PATH = os.path.join(ARTIFACTS_DIR, "obrabotka_baseline.json")

erp = connect_erp()
report = erp.ВнешниеОтчеты.Создать(ERF_PATH, False)

report.НачалоПериода = dt.datetime(2026, 4, 1)
report.ОкончаниеПериода = dt.datetime(2026, 4, 30, 23, 59, 59)

print(f"Запуск АнализРасхождений() для апреля 2026...")
report.АнализРасхождений()

tch = report.ДокументыРасхождения
print(f"ТЧ ДокументыРасхождения: {tch.Количество()} строк")

# Загрузить baseline
with open(BASELINE_PATH, "r", encoding="utf-8") as f:
    baseline = json.load(f)

baseline_rows = baseline["total_rows"]
baseline_sums = baseline["sums_by_podr"]

# Собрать Σ Δ из ТЧ обработки
actual_sums = {}
actual_count = 0
for i in range(tch.Количество()):
    row = tch.Получить(i)
    podr_code = str(row.Подразделение.Код)
    actual_sums[podr_code] = actual_sums.get(podr_code, 0) + float(row.Дельта)
    actual_count += 1

# Сравнить
print("\n=== Acceptance ===")
errors = 0
if actual_count != baseline_rows:
    print(f"FAIL: rows count {actual_count} != baseline {baseline_rows}")
    errors += 1
else:
    print(f"OK: rows count = {actual_count}")

for code, expected in baseline_sums.items():
    actual = actual_sums.get(code, 0)
    diff = abs(actual - expected)
    if diff < 0.01:
        print(f"OK   {code}: Σ Δ = {actual:+.2f} (== baseline)")
    else:
        print(f"FAIL {code}: ожидали {expected:+.2f}, факт {actual:+.2f}, diff {diff:.2f}")
        errors += 1

for code in actual_sums:
    if code not in baseline_sums:
        print(f"FAIL extra: {code} в результате но не в baseline ({actual_sums[code]:+.2f})")
        errors += 1

if errors == 0:
    print("\n*** TEST 1 PASS ***")
    sys.exit(0)
else:
    print(f"\n*** TEST 1 FAIL ({errors} errors) ***")
    sys.exit(1)
```

- [ ] **Step 2: Запустить Test 1**

Run: `python _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_test1_analiz.py`

Expected:
- Печать "ТЧ ДокументыРасхождения: 19 строк"
- 9 строк "OK <код>: Σ Δ = <число>" для каждого подразделения
- "*** TEST 1 PASS ***"
- Exit code 0

⚠️ Если FAIL — STOP, разобраться. Возможные причины: ошибка в FULL OUTER запросе, неверная формула знака, COM не подключился.

- [ ] **Step 3: Commit теста**

```bash
git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_test1_analiz.py
git commit -m "$(cat <<'EOF'
test(balans_klient): TEST 1 PASS — АнализРасхождений найдены 19 виновников

Python COM вызов ВнешниеОтчеты.Создать(ERF) → АнализРасхождений()
для апреля 2026. Сравнение Σ Δ по 9 подразделениям с baseline.
Acceptance: 19 строк, 9 подр, Σ совпали до копейки.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 11: Test 2 — Перепроведение 2 малых документов + проверка РегистраторРасчётов

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_test2_perepro_2docs.py`

Это **критический safety тест** против `feedback_com_repost_skips_registrator_raschetov` — проверяем работает ли перепроведение через COM-обработку для маленьких документов перед массовым прогоном.

- [ ] **Step 1: Написать Тест 2**

Create file `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_test2_perepro_2docs.py`:

```python
# -*- coding: utf-8 -*-
"""
TEST 2 — Перепроведение 2-х малых документов через ПерепровестиДокументы().

Цель: проверить что Записать(Проведение) внутри COM-обработки корректно
создаёт РегистраторРасчётов в РСППС (защита от feedback_com_repost_skips_registrator_raschetov).

Метод:
  1. Анализ → ТЧ заполнена
  2. Очистить ТЧ кроме 2 малых:
     - Списание 000005683 от 29.04.2026 (Логистика, Δ=-780)
     - Списание 000005519 от 24.04.2026 (МД ПРООН, Δ=-1959.32)
  3. Запомнить baseline РегистраторРасчётов для них
  4. Вызвать ПерепровестиДокументы()
  5. Проверить:
     - Обоих Обработан=Истина и НовоеСостояние="ОК"
     - В РСППС РегистраторРасчётов снова есть (не пустой)
     - Δ ≈ 0

Acceptance: 2 документа Обработан=Истина с Δ=0, РегистраторРасчётов сохранён.
"""
import sys, io, datetime as dt
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp

ERF_PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\А_ОбработкаДисбалансаПоСтатьямБаланса.erf"

erp = connect_erp()
report = erp.ВнешниеОтчеты.Создать(ERF_PATH, False)

report.НачалоПериода = dt.datetime(2026, 4, 1)
report.ОкончаниеПериода = dt.datetime(2026, 4, 30, 23, 59, 59)
report.АнализРасхождений()

tch = report.ДокументыРасхождения
print(f"После Анализа: {tch.Количество()} строк ТЧ")

# Оставить только 2 малых
TARGET_NUMBERS = {"000005683", "000005519"}
to_delete = []
for i in range(tch.Количество()):
    row = tch.Получить(i)
    doc_name = str(erp.String(row.Документ)) if row.Документ else ""
    keep = any(num in doc_name for num in TARGET_NUMBERS)
    if not keep:
        to_delete.append(i)
# Удаляем в обратном порядке
for idx in reversed(to_delete):
    tch.Удалить(idx)

print(f"После фильтра: {tch.Количество()} строк (ожидали 2)")
assert tch.Количество() == 2, f"FAIL: ожидали 2 строки, факт {tch.Количество()}"

# Baseline РегистраторРасчётов
def check_registrator_raschetov(doc_ref):
    """Возвращает количество строк в РСППС с этим ДокументРегистратор."""
    q = erp.NewObject("Запрос")
    q.УстановитьПараметр("Док", doc_ref)
    q.Текст = """
    ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Кол
    ИЗ РегистрНакопления.РасчетыСПоставщикамиПоСрокам
    ГДЕ ДокументРегистратор = &Док
    """
    sel = q.Выполнить().Выбрать()
    sel.Следующий()
    return int(sel.Кол)

print("\n=== Baseline РегистраторРасчётов (до перепроведения) ===")
docs_to_check = []
for i in range(tch.Количество()):
    row = tch.Получить(i)
    cnt = check_registrator_raschetov(row.Документ)
    docs_to_check.append((row.Документ, str(erp.String(row.Документ)), cnt))
    print(f"  {str(erp.String(row.Документ))}: РСППС-строк = {cnt}")

# Перепровести
print("\n=== Запуск ПерепровестиДокументы() ===")
report.ПерепровестиДокументы()

# Проверка
print("\n=== Acceptance ===")
errors = 0
for i in range(tch.Количество()):
    row = tch.Получить(i)
    doc_name = str(erp.String(row.Документ))
    status = "OK" if row.Обработан and abs(row.Дельта) < 0.01 else "FAIL"
    print(f"  {status} {doc_name}: Обработан={row.Обработан}, Δ={row.Дельта:+.2f}, состояние='{row.НовоеСостояние}'")
    if row.ОшибкаТекст:
        print(f"    Ошибка: {row.ОшибкаТекст}")
    if status == "FAIL":
        errors += 1
    
    # Проверка РегистраторРасчётов после
    cnt_after = check_registrator_raschetov(row.Документ)
    if cnt_after == 0:
        print(f"  FAIL РегистраторРасчётов: {doc_name} — РСППС пустой после перепроведения!")
        print(f"  ⚠️ КРИТИЧНО: проявилась memory feedback_com_repost_skips_registrator_raschetov")
        errors += 1
    else:
        print(f"  OK РегистраторРасчётов: {doc_name} — РСППС-строк = {cnt_after}")

if errors == 0:
    print("\n*** TEST 2 PASS — перепроведение через COM-обработку работает корректно ***")
    print("*** Можно запускать Test 3 на полную выборку ***")
    sys.exit(0)
else:
    print(f"\n*** TEST 2 FAIL ({errors} errors) — STOP, не запускать Test 3 ***")
    sys.exit(1)
```

- [ ] **Step 2: Запустить Test 2**

Run: `python _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_test2_perepro_2docs.py`

Expected:
- "После фильтра: 2 строки"
- Baseline печать с ненулевым кол-вом РСППС строк
- "*** TEST 2 PASS ***"
- Exit code 0

⚠️ **Если FAIL с сообщением про РегистраторРасчётов — STOP**. Не запускать Test 3. Передать пользователю: «COM-обработка ломает РегистраторРасчётов, нужно перепроводить через UI».

- [ ] **Step 3: Commit теста**

```bash
git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_test2_perepro_2docs.py
git commit -m "$(cat <<'EOF'
test(balans_klient): TEST 2 PASS — перепроведение 2 малых документов через COM ОК

Критическая проверка против feedback_com_repost_skips_registrator_raschetov.
Перепровели 2 малых документа (Логистика 780 + МД ПРООН 1959,32) через
ВнешниеОтчеты.Создать(ERF).ПерепровестиДокументы().
Acceptance: Обработан=Истина, Δ=0, РегистраторРасчётов сохранён.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 12: Test 3 — Полный прогон + перепроведение Свода + верификация плугов

**Files:**
- Create: `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_test3_full_run.py`

- [ ] **Step 1: Написать Тест 3**

Create file `_Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_test3_full_run.py`:

```python
# -*- coding: utf-8 -*-
"""
TEST 3 — Полный прогон обработки + перепроведение Свода + верификация.

После Test 2 (2 малых перепроведены) запускаем массовое перепроведение
ОСТАВШИХСЯ 17 виновников апреля 2026. Затем перепроводим
А_ФинРез_Баланс 00000000004 от 30.04.2026 — плуги должны уйти.

Acceptance:
  - В ТЧ обработки после перепроведения: Обработан=Истина для ≥ 17 строк
  - После перепроведения Свода: в А_ОтчетБаланс_Свод плуги Расхождение=Истина
    по статье ЗПП для 9 подразделений Logistic/Производство/Пьемонт/АВРОРА/
    Кривий Ріг/МД ПРООН/МАЗ/Экскаватор/Телескоп = 0 (Σ |КО| < 0.01)
"""
import sys, io, datetime as dt
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)
from _common import connect_erp

ERF_PATH = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\Обработки\А_ОбработкаДисбалансаПоСтатьямБаланса.erf"

erp = connect_erp()
report = erp.ВнешниеОтчеты.Создать(ERF_PATH, False)
report.НачалоПериода = dt.datetime(2026, 4, 1)
report.ОкончаниеПериода = dt.datetime(2026, 4, 30, 23, 59, 59)

print("=== Шаг 1: Анализ ===")
report.АнализРасхождений()
tch = report.ДокументыРасхождения
print(f"ТЧ: {tch.Количество()} строк")
n_before = tch.Количество()

print("\n=== Шаг 2: Массовое перепроведение ===")
report.ПерепровестиДокументы()

# Подсчитать результаты
ok = 0
err = 0
ostalas = 0
errors_detail = []
for i in range(tch.Количество()):
    row = tch.Получить(i)
    if row.Обработан:
        ok += 1
    elif row.ОшибкаТекст:
        err += 1
        errors_detail.append(f"{str(erp.String(row.Документ))}: {row.ОшибкаТекст}")
    else:
        ostalas += 1

print(f"\nРезультат перепроведения:")
print(f"  OK: {ok}")
print(f"  Осталась Δ: {ostalas}")
print(f"  Ошибок: {err}")
if errors_detail:
    print("\nДетали ошибок:")
    for e in errors_detail[:5]:
        print(f"  {e}")

print("\n=== Шаг 3: Перепроведение А_ФинРез_Баланс 00000000004 от 30.04.2026 ===")
q = erp.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ ПЕРВЫЕ 1 Д.Ссылка
ИЗ Документ.А_ФинРез_Баланс КАК Д
ГДЕ Д.Номер = "00000000004" И Д.Дата МЕЖДУ ДАТАВРЕМЯ(2026,4,30) И ДАТАВРЕМЯ(2026,4,30,23,59,59)
    И Д.Проведен И НЕ Д.ПометкаУдаления
"""
sel = q.Выполнить().Выбрать()
if sel.Следующий():
    svod_ref = sel.Ссылка
    svod_obj = svod_ref.ПолучитьОбъект()
    svod_obj.Записать(erp.РежимЗаписиДокумента.Проведение)
    print(f"  Перепроведено: {erp.String(svod_ref)}")
else:
    print("  ⚠️ А_ФинРез_Баланс 00000000004 от 30.04 не найден — пропускаем")

print("\n=== Шаг 4: Верификация плугов в А_ОтчетБаланс_Свод ===")
q = erp.NewObject("Запрос")
q.Текст = """
ВЫБРАТЬ
    Т.Подразделение.Код КАК ПодрКод,
    ПРЕДСТАВЛЕНИЕ(Т.Подразделение) КАК ПодрИмя,
    Т.СуммаКонечныйОстаток КАК КО
ИЗ РегистрСведений.А_ОтчетБаланс_Свод КАК Т
ГДЕ Т.Расхождение = ИСТИНА
    И Т.Регистратор.Месяц = ДАТАВРЕМЯ(2026,4,1)
    И Т.Статья = ЗНАЧЕНИЕ(ПланВидовХарактеристик.СтатьиАктивовПассивов.ЗадолженностьПередПоставщиками)
    И Т.Source = ЗНАЧЕНИЕ(Перечисление.ИсточникиУправленческогоБаланса.РасчетыСПоставщикамиПоСрокам)
    И Т.Подразделение.Код В ("00-001029","00-001022","00-001026","0Ц-000004","00-001003","0Ц-000001","00-000210","00-000219","00-000190")
"""
r = q.Выполнить().Выгрузить()
sum_abs_left = 0
for i in range(r.Количество()):
    row = r.Получить(i)
    ко = float(row.КО)
    if abs(ко) >= 0.01:
        print(f"  Ещё плуг: {row.ПодрИмя} ({row.ПодрКод}): КО={ко:+.2f}")
    sum_abs_left += abs(ко)

print(f"\nΣ |КО| оставшихся плугов: {sum_abs_left:.2f}")

print("\n=== Acceptance ===")
errors = 0
if ok + 2 >= n_before:  # +2 из Test 2 уже Обработаны
    print(f"OK: перепроведено {ok}/{n_before} (acceptance ≥ {n_before-2})")
else:
    print(f"FAIL: перепроведено только {ok}/{n_before}")
    errors += 1

if sum_abs_left < 0.01:
    print("OK: плуги Расхождение=Истина для 9 подразделений ушли (Σ |КО| < 0.01)")
else:
    print(f"FAIL: остались плуги на Σ {sum_abs_left:.2f}")
    errors += 1

if errors == 0:
    print("\n*** TEST 3 PASS — обработка работает end-to-end ***")
    sys.exit(0)
else:
    print(f"\n*** TEST 3 FAIL ({errors} errors) ***")
    sys.exit(1)
```

- [ ] **Step 2: Запустить Test 3**

Run: `python _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_test3_full_run.py`

Expected:
- "OK: перепроведено 17/17" (или близкое)
- "OK: плуги ушли"
- "*** TEST 3 PASS ***"

⚠️ Если осталась Δ в каких-то документах — это нормальный исход (некоторые могут требовать перепроведения источника долга — ПриобретениеТоваровУслуг — отдельно). Сообщить пользователю список.

- [ ] **Step 3: Commit финального теста**

```bash
git add _Rarzrabotki/notebook/knowledge_Balanse_klient/Python/test/obrabotka_test3_full_run.py
git commit -m "$(cat <<'EOF'
test(balans_klient): TEST 3 — полный E2E прогон obrabotka_disbalansa_statyam

Массовое перепроведение всех виновников апреля 2026 + перепроведение
А_ФинРез_Баланс 00000000004 + верификация плугов Расхождение=Истина
по 9 подразделениям статья ЗПП.

Acceptance: Σ |КО| оставшихся плугов < 0.01 для целевых 9 подр.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 13: Финальная сводка и обновление knowledge

**Files:**
- Modify: `_Rarzrabotki/notebook/knowledge_Balanse_klient/FINDINGS.md`
- Modify: `_Rarzrabotki/notebook/knowledge_Balanse_klient/README.md`

- [ ] **Step 1: Добавить раздел в FINDINGS.md**

Append к `_Rarzrabotki/notebook/knowledge_Balanse_klient/FINDINGS.md` (в самом верху, после статуса):

```markdown

## 🎯 Phase 7 (2026-05-25): Автоматизация обработки расхождений — реализовано

Создана обработка `А_ОбработкаДисбалансаПоСтатьямБаланса.erf`:
- Автоматически находит первичные документы-виновники через FULL OUTER ПАП vs РСКПС/РСППС по `ДокументРегистратор`
- Перепроводит их в транзакциях per-документ (`Записать(Проведение)`)
- Пересчитывает Δ после перепроведения, статус виден в ТЧ

**E2E test workflow:** 3 Python COM теста (`obrabotka_test1/2/3_*.py`). Test 2 — критическая проверка `feedback_com_repost_skips_registrator_raschetov` на 2 малых документах перед массовым прогоном.

**Spec/Plan:** `docs/superpowers/specs/2026-05-25-obrabotka-disbalansa-statyam-design.md` + `docs/superpowers/plans/2026-05-25-obrabotka-disbalansa-statyam-plan.md`

**Применение:** для любого периода / любых подразделений с плугами по РСКПС/РСППС в `А_ОтчетБаланс_Свод`. На декабрь 2025 — применять с осторожностью (есть отдельный фикс DvAktPas FIX-2026-05-23 для случая ПереносАванса).
```

- [ ] **Step 2: Обновить README.md — добавить новый артефакт**

В разделе «Главный entry-point для ИИ-сессий» добавить:

```markdown
- **Обработка обнаружения и устранения**: `_Rarzrabotki/Обработки/А_ОбработкаДисбалансаПоСтатьямБаланса.erf`
  - Запускать из 1С UI или Python COM через `ВнешниеОтчеты.Создать()`
  - Дизайн: [docs/superpowers/specs/2026-05-25-obrabotka-disbalansa-statyam-design.md](../../docs/superpowers/specs/2026-05-25-obrabotka-disbalansa-statyam-design.md)
```

- [ ] **Step 3: Commit knowledge**

```bash
git add _Rarzrabotki/notebook/knowledge_Balanse_klient/FINDINGS.md \
        _Rarzrabotki/notebook/knowledge_Balanse_klient/README.md
git commit -m "$(cat <<'EOF'
docs(balans_klient): Phase 7 — добавлена обработка автоматического устранения

Обновлены FINDINGS.md и README.md с информацией об автоматизированной
обработке. Готова для применения на любом периоде с плугами РСКПС/РСППС.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Acceptance criteria (рекап)

После выполнения всех 13 задач:

- [x] **Task 1**: Baseline JSON с 19 виновниками зафиксирован
- [x] **Task 2-9**: Обработка собрана, ERF файл создан
- [x] **Task 10 (Test 1)**: АнализРасхождений → ТЧ совпадает с baseline до копейки
- [x] **Task 11 (Test 2)**: 2 малых документа перепроведены, РегистраторРасчётов сохранён
- [x] **Task 12 (Test 3)**: Полный прогон + Свод → плуги ушли (Σ |КО| < 0.01)
- [x] **Task 13**: Knowledge обновлён

**Σ-инвариант штатного отчёта `Отчёт.УправленческийБаланс`** не должен ухудшиться — проверить после Test 3 если возникнут сомнения (запросом по всем 9 подразделениям что |Актив|=|Пассив| per организация).

## Risks & mitigation (рекап из spec)

| Risk | Mitigation в плане |
|---|---|
| COM-обработка ломает РегистраторРасчётов | Test 2 на 2-х документах перед массовым прогоном |
| FULL OUTER падает на типах | Pre-test через MCP в Task 3 + Python COM в Task 10 |
| Δ не уходит после перепроведения (нужна перепроведения источника долга) | Видно в колонке НовоеСостояние = "Δ=<новое>", пользователь решает |
| Битые ссылки документов | Перехват в `Попытка`, текст в ОшибкаТекст |

## Memory checklist

- ✅ `feedback_query_pre_check` + Rule #-1 — Task 3 pre-test SQL через MCP
- ✅ `feedback_com_repost_skips_registrator_raschetov` — Task 11 (Test 2) safety check
- ✅ `feedback_no_doc_delete_in_tests` — только Записать(Проведение), без Удалить()
- ✅ `feedback_no_typical_register_changes` — типовые регистры только читаем
- ✅ `feedback_balans_etalon_period_serverside` — в BSL даты передаются как параметры (Запрос.УстановитьПараметр), при необходимости в подзапросах используется НАЧАЛОПЕРИОДА/КОНЕЦПЕРИОДА серверно
- ✅ `feedback_designer_cache_invalidation` — НЕ применимо (ERF читается из файла каждый раз)
- ✅ `feedback_balans_klient_perenosa_avansa_only` — НЕ применяется (наш случай не ПереносАванса)

---

**End of plan.**
