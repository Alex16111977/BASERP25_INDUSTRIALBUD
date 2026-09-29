#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область ПрограммныйИнтерфейс

// Книга Excel (xlsx) з моделі аркуша фінзвіту (МодельКнигиExcel у модулі об'єкта): значення разом
// з формулами (кеш значення + формула, перерахунок при відкритті), стилі директорського звіту,
// групування рядків, закріплення шапки, примітки з розміром вікна під текст.
// Пишеться напряму в Office Open XML: ТабличныйДокумент.Записать(XLSX) формул не пише,
// закріплення і розмір приміток не переносить (перевірено на 8.3.20).
//
// Модель — Структура:
//   ИмяЛиста    — Строка;
//   Ширины      — Массив строк-чисел (ширина колонок A, B, …);
//   Строки      — Массив Структур (Высота, Скрыта, Уровень, Ячейки); номер рядка = індекс + 1;
//                 Ячейки — Массив Структур (Колонка — номер, Значение — Строка/Число,
//                 Формула — Строка без «=», Стиль — індекс cellXfs з ТекстСтилейКниги);
//   Объединения — Массив адрес «A7:B7»;
//   Примечания  — Массив Структур (Строка, Колонка, Текст).
Функция ДвоичныеДанныеКнигиExcel(Модель) Экспорт

	Разделитель = ПолучитьРазделительПути();
	Каталог = ПолучитьИмяВременногоФайла("xlsxparts");
	СоздатьКаталог(Каталог);
	ЕстьПримечания = Модель.Примечания.Количество() > 0;

	ЗаписатьЧастьКниги(Каталог, "[Content_Types].xml", ТекстТиповСодержимого(ЕстьПримечания));
	ЗаписатьЧастьКниги(Каталог, "_rels/.rels",
		"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		+ "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">"
		+ "<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>"
		+ "</Relationships>");
	ЗаписатьЧастьКниги(Каталог, "xl/workbook.xml",
		"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		+ "<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" "
		+ "xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">"
		+ "<bookViews><workbookView/></bookViews>"
		+ "<sheets><sheet name=""" + ЭкранироватьXML(Модель.ИмяЛиста) + """ sheetId=""1"" r:id=""rId1""/></sheets>"
		+ "<calcPr calcId=""191029"" fullCalcOnLoad=""1""/>"
		+ "</workbook>");
	ЗаписатьЧастьКниги(Каталог, "xl/_rels/workbook.xml.rels",
		"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		+ "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">"
		+ "<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>"
		+ "<Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>"
		+ "</Relationships>");
	ЗаписатьЧастьКниги(Каталог, "xl/styles.xml", ТекстСтилейКниги());
	ЗаписатьЧастьКниги(Каталог, "xl/worksheets/sheet1.xml", ТекстЛиста(Модель, ЕстьПримечания));
	Если ЕстьПримечания Тогда
		ЗаписатьЧастьКниги(Каталог, "xl/worksheets/_rels/sheet1.xml.rels",
			"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
			+ "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">"
			+ "<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"" Target=""../comments1.xml""/>"
			+ "<Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"" Target=""../drawings/vmlDrawing1.vml""/>"
			+ "</Relationships>");
		ЗаписатьЧастьКниги(Каталог, "xl/comments1.xml", ТекстПримечаний(Модель));
		ЗаписатьЧастьКниги(Каталог, "xl/drawings/vmlDrawing1.vml", ТекстVMLПримечаний(Модель));
	КонецЕсли;

	ИмяФайла = ПолучитьИмяВременногоФайла("xlsx");
	Архив = Новый ЗаписьZipФайла(ИмяФайла);
	Архив.Добавить(Каталог + Разделитель + "*", РежимСохраненияПутейZIP.СохранятьОтносительныеПути,
		РежимОбработкиПодкаталоговZIP.ОбрабатыватьРекурсивно);
	Архив.Записать();

	Результат = Новый ДвоичныеДанные(ИмяФайла);
	УдалитьФайлы(Каталог);
	УдалитьФайлы(ИмяФайла);
	Возврат Результат;

КонецФункции

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

Процедура ЗаписатьЧастьКниги(Каталог, Путь, Текст)

	Разделитель = ПолучитьРазделительПути();
	ПолныйПуть = Каталог + Разделитель + СтрЗаменить(Путь, "/", Разделитель);
	ФайлЧасти = Новый Файл(ПолныйПуть);
	КаталогЧасти = Новый Файл(ФайлЧасти.Путь);
	Если НЕ КаталогЧасти.Существует() Тогда
		СоздатьКаталог(КаталогЧасти.ПолноеИмя);
	КонецЕсли;
	// UTF-8 без BOM — як пише Excel
	ПолучитьДвоичныеДанныеИзСтроки(Текст, КодировкаТекста.UTF8, Ложь).Записать(ПолныйПуть);

КонецПроцедуры

Функция ТекстТиповСодержимого(ЕстьПримечания)

	Текст = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		+ "<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">"
		+ "<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>"
		+ "<Default Extension=""xml"" ContentType=""application/xml""/>"
		+ "<Default Extension=""vml"" ContentType=""application/vnd.openxmlformats-officedocument.vmlDrawing""/>"
		+ "<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>"
		+ "<Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>"
		+ "<Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>";
	Если ЕстьПримечания Тогда
		Текст = Текст + "<Override PartName=""/xl/comments1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml""/>";
	КонецЕсли;
	Возврат Текст + "</Types>";

КонецФункции

// Стилі (індекси cellXfs — ті, що ставить модель):
//   1 шапка Arial 10 ж, 2 шапка Arial 9 ж, 3 шапка G5, 4 назва об'єкта, 5 число об'єкта,
//   6 текст рядка суми, 7 число рядка суми, 8 текст договору, 9 число договору (нуль — порожньо),
//   10 текст «Итого», 11 число «Итого». Кольори — з директорського звіту 25.09.2026.
Функция ТекстСтилейКниги()

	Рамка = "<left style=""thin""><color auto=""1""/></left><right style=""thin""><color auto=""1""/></right>"
		+ "<top style=""thin""><color auto=""1""/></top><bottom style=""thin""><color auto=""1""/></bottom><diagonal/>";
	Центр = "<alignment horizontal=""center"" vertical=""center"" wrapText=""1""/>";

	Возврат "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		+ "<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">"
		+ "<numFmts count=""2"">"
		+ "<numFmt numFmtId=""164"" formatCode=""#,##0.00_ ;[Red]\-#,##0.00\ ""/>"
		+ "<numFmt numFmtId=""165"" formatCode=""#,##0.00_ ;[Red]\-#,##0.00\ ;""/>"
		+ "</numFmts>"
		+ "<fonts count=""6"">"
		+ "<font><sz val=""10""/><name val=""Arial""/><family val=""2""/></font>"
		+ "<font><b/><sz val=""10""/><color rgb=""FF003F2F""/><name val=""Arial""/><family val=""2""/></font>"
		+ "<font><b/><sz val=""9""/><color rgb=""FF003F2F""/><name val=""Arial""/><family val=""2""/></font>"
		+ "<font><b/><sz val=""11""/><color rgb=""FF003F2F""/><name val=""Arial""/><family val=""2""/></font>"
		+ "<font><sz val=""8""/><name val=""Arial""/><family val=""2""/></font>"
		+ "<font><b/><sz val=""11""/><name val=""Arial""/><family val=""2""/></font>"
		+ "</fonts>"
		+ "<fills count=""6"">"
		+ "<fill><patternFill patternType=""none""/></fill>"
		+ "<fill><patternFill patternType=""gray125""/></fill>"
		+ "<fill><patternFill patternType=""solid""><fgColor rgb=""FFC3D69B""/><bgColor indexed=""64""/></patternFill></fill>"
		+ "<fill><patternFill patternType=""solid""><fgColor rgb=""FFD7E4BD""/><bgColor indexed=""64""/></patternFill></fill>"
		+ "<fill><patternFill patternType=""solid""><fgColor rgb=""FF92CDDC""/><bgColor indexed=""64""/></patternFill></fill>"
		+ "<fill><patternFill patternType=""solid""><fgColor rgb=""FFD6E5CB""/><bgColor indexed=""64""/></patternFill></fill>"
		+ "</fills>"
		+ "<borders count=""2""><border><left/><right/><top/><bottom/><diagonal/></border><border>" + Рамка + "</border></borders>"
		+ "<cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs>"
		+ "<cellXfs count=""12"">"
		+ "<xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/>"
		+ "<xf numFmtId=""0"" fontId=""1"" fillId=""2"" borderId=""1"" xfId=""0"" applyFont=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"">" + Центр + "</xf>"
		+ "<xf numFmtId=""0"" fontId=""2"" fillId=""2"" borderId=""1"" xfId=""0"" applyFont=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"">" + Центр + "</xf>"
		+ "<xf numFmtId=""0"" fontId=""2"" fillId=""3"" borderId=""1"" xfId=""0"" applyFont=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"">" + Центр + "</xf>"
		+ "<xf numFmtId=""0"" fontId=""3"" fillId=""4"" borderId=""1"" xfId=""0"" applyFont=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1""><alignment vertical=""center"" wrapText=""1""/></xf>"
		+ "<xf numFmtId=""164"" fontId=""0"" fillId=""4"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"">" + Центр + "</xf>"
		+ "<xf numFmtId=""0"" fontId=""4"" fillId=""3"" borderId=""1"" xfId=""0"" applyFont=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1""><alignment vertical=""top"" wrapText=""1""/></xf>"
		+ "<xf numFmtId=""164"" fontId=""4"" fillId=""3"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFont=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"">" + Центр + "</xf>"
		+ "<xf numFmtId=""0"" fontId=""4"" fillId=""0"" borderId=""1"" xfId=""0"" applyFont=""1"" applyBorder=""1"" applyAlignment=""1""><alignment vertical=""top"" wrapText=""1""/></xf>"
		+ "<xf numFmtId=""165"" fontId=""4"" fillId=""0"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFont=""1"" applyBorder=""1"" applyAlignment=""1"">" + Центр + "</xf>"
		+ "<xf numFmtId=""0"" fontId=""5"" fillId=""5"" borderId=""1"" xfId=""0"" applyFont=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1""><alignment vertical=""center""/></xf>"
		+ "<xf numFmtId=""4"" fontId=""5"" fillId=""5"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFont=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"">" + Центр + "</xf>"
		+ "</cellXfs>"
		+ "<cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles>"
		+ "</styleSheet>";

КонецФункции

Функция ТекстЛиста(Модель, ЕстьПримечания)

	Части = Новый Массив;
	Части.Добавить("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		+ "<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" "
		+ "xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">"
		+ "<sheetPr><outlinePr summaryBelow=""1""/></sheetPr>"
		+ "<sheetViews><sheetView workbookViewId=""0"">"
		+ "<pane ySplit=""5"" topLeftCell=""A6"" activePane=""bottomLeft"" state=""frozen""/>"
		+ "<selection pane=""bottomLeft"" activeCell=""A6"" sqref=""A6""/>"
		+ "</sheetView></sheetViews>"
		+ "<sheetFormatPr defaultRowHeight=""11.25"" outlineLevelRow=""1""/>");

	Части.Добавить("<cols>");
	НомерКолонки = 0;
	Для Каждого Ширина Из Модель.Ширины Цикл
		НомерКолонки = НомерКолонки + 1;
		Н = ЧислоXML(НомерКолонки);
		Части.Добавить("<col min=""" + Н + """ max=""" + Н + """ width=""" + Ширина + """ customWidth=""1""/>");
	КонецЦикла;
	Части.Добавить("</cols><sheetData>");

	НомерСтроки = 0;
	Для Каждого СтрокаМодели Из Модель.Строки Цикл
		НомерСтроки = НомерСтроки + 1;
		Н = ЧислоXML(НомерСтроки);
		Атрибуты = " r=""" + Н + """";
		Если СтрокаМодели.Высота > 0 Тогда
			Атрибуты = Атрибуты + " ht=""" + XMLСтрока(СтрокаМодели.Высота) + """ customHeight=""1""";
		КонецЕсли;
		Если СтрокаМодели.Скрыта Тогда
			Атрибуты = Атрибуты + " hidden=""1""";
		КонецЕсли;
		Если СтрокаМодели.Уровень > 0 Тогда
			Атрибуты = Атрибуты + " outlineLevel=""" + ЧислоXML(СтрокаМодели.Уровень) + """";
		КонецЕсли;
		Если СтрокаМодели.Ячейки.Количество() = 0 Тогда
			Части.Добавить("<row" + Атрибуты + "/>");
			Продолжить;
		КонецЕсли;
		Части.Добавить("<row" + Атрибуты + ">");
		Для Каждого Ячейка Из СтрокаМодели.Ячейки Цикл
			Части.Добавить(ТекстЯчейки(БукваКолонки(Ячейка.Колонка) + Н, Ячейка));
		КонецЦикла;
		Части.Добавить("</row>");
	КонецЦикла;
	Части.Добавить("</sheetData>");

	Если Модель.Объединения.Количество() > 0 Тогда
		Части.Добавить("<mergeCells count=""" + ЧислоXML(Модель.Объединения.Количество()) + """>");
		Для Каждого Диапазон Из Модель.Объединения Цикл
			Части.Добавить("<mergeCell ref=""" + Диапазон + """/>");
		КонецЦикла;
		Части.Добавить("</mergeCells>");
	КонецЕсли;
	Части.Добавить("<pageMargins left=""0.7"" right=""0.7"" top=""0.75"" bottom=""0.75"" header=""0.3"" footer=""0.3""/>"
		+ "<pageSetup paperSize=""9"" orientation=""portrait""/>");
	Если ЕстьПримечания Тогда
		Части.Добавить("<legacyDrawing r:id=""rId2""/>");
	КонецЕсли;
	Части.Добавить("</worksheet>");

	Возврат СтрСоединить(Части, "");

КонецФункции

// Рядок — inlineStr; число — значення, з формулою — формула і кеш значення. Нуль без формули
// не пишеться (порожня комірка зі стилем), як у директорському звіті.
Функция ТекстЯчейки(Адрес, Ячейка)

	Начало = "<c r=""" + Адрес + """ s=""" + ЧислоXML(Ячейка.Стиль) + """";
	Если ТипЗнч(Ячейка.Значение) = Тип("Строка") Тогда
		Если ПустаяСтрока(Ячейка.Значение) Тогда
			Возврат Начало + "/>";
		КонецЕсли;
		Возврат Начало + " t=""inlineStr""><is><t xml:space=""preserve"">" + ЭкранироватьXML(Ячейка.Значение)
			+ "</t></is></c>";
	КонецЕсли;
	Если ПустаяСтрока(Ячейка.Формула) Тогда
		Если Ячейка.Значение = 0 Тогда
			Возврат Начало + "/>";
		КонецЕсли;
		Возврат Начало + "><v>" + XMLСтрока(Ячейка.Значение) + "</v></c>";
	КонецЕсли;
	Возврат Начало + "><f>" + ЭкранироватьXML(Ячейка.Формула) + "</f><v>" + XMLСтрока(Ячейка.Значение) + "</v></c>";

КонецФункции

Функция ТекстПримечаний(Модель)

	Части = Новый Массив;
	Части.Добавить("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
		+ "<comments xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">"
		+ "<authors><author>1С</author></authors><commentList>");
	Для Каждого Примечание Из Модель.Примечания Цикл
		Части.Добавить("<comment ref=""" + БукваКолонки(Примечание.Колонка) + ЧислоXML(Примечание.Строка)
			+ """ authorId=""0""><text><r><rPr><sz val=""8""/><rFont val=""Tahoma""/><family val=""2""/></rPr>"
			+ "<t xml:space=""preserve"">" + ЭкранироватьXML(Примечание.Текст) + "</t></r></text></comment>");
	КонецЦикла;
	Части.Добавить("</commentList></comments>");
	Возврат СтрСоединить(Части, "");

КонецФункции

// Вікно примітки — під текст: висота = рядки тексту, ширина = найдовший рядок (Tahoma 8 ≈ 4,6 pt
// на символ). Excel бере розмір з x:Anchor (колонки/рядки), style — для сумісності.
Функция ТекстVMLПримечаний(Модель)

	Части = Новый Массив;
	Части.Добавить("<xml xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"" "
		+ "xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
		+ "<o:shapelayout v:ext=""edit""><o:idmap v:ext=""edit"" data=""1""/></o:shapelayout>"
		+ "<v:shapetype id=""_x0000_t202"" coordsize=""21600,21600"" o:spt=""202"" path=""m,l,21600r21600,l21600,xe"">"
		+ "<v:stroke joinstyle=""miter""/><v:path gradientshapeok=""t"" o:connecttype=""rect""/></v:shapetype>");
	НомерФигуры = 1024;
	Для Каждого Примечание Из Модель.Примечания Цикл
		НомерФигуры = НомерФигуры + 1;
		Строки = СтрРазделить(Примечание.Текст, Символы.ПС, Истина);
		МаксДлина = 10;
		Для Каждого СтрокаТекста Из Строки Цикл
			МаксДлина = Макс(МаксДлина, СтрДлина(СтрокаТекста));
		КонецЦикла;
		ШиринаPt = Мин(620, Макс(150, МаксДлина * 4.6 + 12));
		ВысотаPt = Строки.Количество() * 11 + 8;
		КолонокШирины = Цел(ШиринаPt / 96) + 1;       // колонка 18,5 ≈ 96 pt
		РядковВисоты = Цел(ВысотаPt / 11.25) + 1;      // рядок за замовчуванням 11,25 pt
		Колонка0 = Примечание.Колонка - 1;
		Строка0 = Примечание.Строка - 1;
		Якорь = ЧислоXML(Колонка0 + 1) + ", 15, " + ЧислоXML(Макс(0, Строка0 - 1)) + ", 2, "
			+ ЧислоXML(Колонка0 + 1 + КолонокШирины) + ", 15, " + ЧислоXML(Макс(0, Строка0 - 1) + РядковВисоты) + ", 2";
		Части.Добавить("<v:shape id=""_x0000_s" + ЧислоXML(НомерФигуры) + """ type=""#_x0000_t202"" "
			+ "style=""position:absolute;margin-left:80pt;margin-top:2pt;width:" + XMLСтрока(Окр(ШиринаPt, 1))
			+ "pt;height:" + XMLСтрока(Окр(ВысотаPt, 1)) + "pt;z-index:" + ЧислоXML(НомерФигуры - 1024)
			+ ";visibility:hidden"" fillcolor=""#ffffe1"" o:insetmode=""auto"">"
			+ "<v:fill color2=""#ffffe1""/><v:shadow on=""t"" color=""black"" obscured=""t""/>"
			+ "<v:path o:connecttype=""none""/><v:textbox style=""mso-direction-alt:auto""><div style=""text-align:left""></div></v:textbox>"
			+ "<x:ClientData ObjectType=""Note""><x:MoveWithCells/><x:SizeWithCells/>"
			+ "<x:Anchor>" + Якорь + "</x:Anchor><x:AutoFill>False</x:AutoFill>"
			+ "<x:Row>" + ЧислоXML(Строка0) + "</x:Row><x:Column>" + ЧислоXML(Колонка0) + "</x:Column>"
			+ "</x:ClientData></v:shape>");
	КонецЦикла;
	Части.Добавить("</xml>");
	Возврат СтрСоединить(Части, "");

КонецФункции

Функция БукваКолонки(Номер)

	Возврат Сред("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Номер, 1);

КонецФункции

Функция ЧислоXML(Число)

	Возврат Формат(Число, "ЧН=0; ЧГ=0");

КонецФункции

// Екранування для XML + прибирання керівних символів, недопустимих у XML 1.0
// (у найменуваннях договорів трапляються).
Функция ЭкранироватьXML(Знач Текст)

	Текст = СтрЗаменить(Текст, "&", "&amp;");
	Текст = СтрЗаменить(Текст, "<", "&lt;");
	Текст = СтрЗаменить(Текст, ">", "&gt;");
	Текст = СтрЗаменить(Текст, """", "&quot;");
	Текст = СтрЗаменить(Текст, Символы.ВК + Символы.ПС, Символы.ПС);
	Для Код = 1 По 31 Цикл
		Если Код <> 9 И Код <> 10 И Код <> 13 Тогда
			Текст = СтрЗаменить(Текст, Символ(Код), " ");
		КонецЕсли;
	КонецЦикла;
	Возврат Текст;

КонецФункции

#КонецОбласти

#КонецЕсли
