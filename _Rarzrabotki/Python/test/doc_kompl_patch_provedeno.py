# -*- coding: utf-8 -*-
"""Патч маніфесту Документ.РасчетКомплектаций: колонки «Проведено» у ТЧ документів.

Додає по два реквізити Булево в ТЧ ДокументиКомплектації та ДокументиМалоценки:
    ПроведенПоНормам, ПроведенДодаткова
Це знімок стану цільових документів (1С не дає обчислюваних колонок у таблиці,
привʼязаній до ТЧ обʼєкта) — оновлюється ОбновитьСостояниеДокументов().

Ідемпотентний: якщо реквізити вже є — нічого не робить.
Спека: docs/superpowers/specs/2026-08-04-podmenyu-dokumenty-design.md
"""
import io
import re
import sys
import uuid

sys.stdout.reconfigure(encoding='utf-8')

ФАЙЛ = (r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh"
        r"\Documents\РасчетКомплектаций.xml")

НОВЫЕ = [
    ("ПроведенПоНормам", "Проведений «По нормам»"),
    ("ПроведенДодаткова", "Проведений «Понад норму»"),
]
ТЧ_ЦЕЛИ = ("ДокументиКомплектації", "ДокументиМалоценки")


def реквизит_булево(имя, синоним):
    """Реквизит ТЧ типа Булево. Отступы — 6 табов, как у соседних в ЭТОМ файле."""
    о = "\t" * 5
    return f"""{о}<Attribute uuid="{uuid.uuid4()}">
{о}\t<Properties>
{о}\t\t<Name>{имя}</Name>
{о}\t\t<Synonym>
{о}\t\t\t<v8:item>
{о}\t\t\t\t<v8:lang>uk</v8:lang>
{о}\t\t\t\t<v8:content>{синоним}</v8:content>
{о}\t\t\t</v8:item>
{о}\t\t</Synonym>
{о}\t\t<Comment/>
{о}\t\t<Type>
{о}\t\t\t<v8:Type>xs:boolean</v8:Type>
{о}\t\t</Type>
{о}\t\t<PasswordMode>false</PasswordMode>
{о}\t\t<Format/>
{о}\t\t<EditFormat/>
{о}\t\t<ToolTip/>
{о}\t\t<MarkNegatives>false</MarkNegatives>
{о}\t\t<Mask/>
{о}\t\t<MultiLine>false</MultiLine>
{о}\t\t<ExtendedEdit>false</ExtendedEdit>
{о}\t\t<MinValue xsi:nil="true"/>
{о}\t\t<MaxValue xsi:nil="true"/>
{о}\t\t<FillChecking>DontCheck</FillChecking>
{о}\t\t<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
{о}\t\t<ChoiceParameterLinks/>
{о}\t\t<ChoiceParameters/>
{о}\t\t<QuickChoice>Auto</QuickChoice>
{о}\t\t<CreateOnInput>Auto</CreateOnInput>
{о}\t\t<ChoiceForm/>
{о}\t\t<LinkByType/>
{о}\t\t<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
{о}\t\t<Indexing>DontIndex</Indexing>
{о}\t\t<FullTextSearch>Use</FullTextSearch>
{о}\t\t<DataHistory>Use</DataHistory>
{о}\t</Properties>
{о}</Attribute>"""


def main():
    исходный = io.open(ФАЙЛ, encoding='utf-8-sig').read()
    s = исходный
    добавлено = 0

    for тч in ТЧ_ЦЕЛИ:
        # блок конкретной ТЧ: от <TabularSection ...> с нужным именем до </TabularSection>
        шаблон = (r'(<TabularSection uuid="[^"]+">\s*<InternalInfo>(?:(?!</TabularSection>).)*?'
                  r'<Name>' + re.escape(тч) + r'</Name>(?:(?!</TabularSection>).)*?)'
                  r'(\n\t{4}</ChildObjects>\n\t{3}</TabularSection>)')
        m = re.search(шаблон, s, re.S)
        if m is None:
            print(f"  ⚠ ТЧ {тч} не знайдена — пропуск")
            continue

        блок = m.group(0)
        вставка = []
        for имя, синоним in НОВЫЕ:
            if re.search(r'<Name>' + re.escape(имя) + r'</Name>', блок):
                print(f"  {тч}.{имя}: вже є — пропуск")
                continue
            вставка.append(реквизит_булево(имя, синоним))
            добавлено += 1
            print(f"  {тч}.{имя}: додано")

        если_есть = "\n" + "\n".join(вставка) if вставка else ""
        s = s[:m.start()] + m.group(1) + если_есть + m.group(2) + s[m.end():]

    if добавлено == 0:
        print("\nЗмін немає — файл не переписано.")
        return

    # 1С зберігає цей файл із LF; зберігаємо як є, щоб не плодити diff (§Д17)
    with io.open(ФАЙЛ, "w", encoding='utf-8-sig', newline="\n") as f:
        f.write(s)
    print(f"\nЗаписано: {ФАЙЛ}  (+{добавлено} реквізитів)")


if __name__ == "__main__":
    main()
