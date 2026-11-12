let
    // Параметры, используемые в запросе
    filePart_1 = "demo-consumable-movement",
    sheetName_1 = "DEMO-CONSUMABLE-MOVEMENT",
    headerMarker_1 = "Наименование",
    endMarker_1 = "диспетчер",

    filePart_2 = "demo-packaging-stock-b",
    sheetName_2 = "Остатки",
    headerMarker_2 = "№ п/п",
    rowsToSkip_2 = 1,

    // Пользовательская функция на безопасное получение файла
    GetFileContent = (folderTable as table, fileNamePart as text) as binary =>
		let
			filtered = Table.SelectRows(folderTable, each Text.Contains(Text.Lower([Name]), Text.Lower(fileNamePart))),
            result = if Table.RowCount(filtered) > 0 then filtered{0}[Content] else error "Файл не найден"
        in
            result,

	// Пользовательская функция на безопасное получение листа
	GetSheetData = (workbook as table, sheetName as text) as table =>
        let
            filtered = Table.SelectRows(workbook, each Text.Contains([Name], sheetName) and [Kind] = "Sheet"),
            result = if Table.RowCount(filtered) > 0 then filtered{0}[Data] else error "Лист не найден"
        in
            result,

	// Обработка источника - DEMO-CONSUMABLE-MOVEMENT
    consumable_1 = let
		fileContent = GetFileContent(FolderContent, filePart_1),
        workbook = Excel.Workbook(fileContent),
        rawData = GetSheetData(workbook, sheetName_1),

		// Поиск границ диапазона данных
		startPos = List.PositionOf(rawData[Column2], headerMarker_1),
		endPos = List.PositionOf(List.Transform(rawData[Column2], each Text.Contains(Text.Lower(_), endMarker_1)), true),

		// Поиск границ диапазона данных
		validRange = startPos >= 0 and endPos >= 0 and endPos > startPos,
		dataRange = if validRange then Table.FirstN(Table.Skip(rawData, startPos), endPos - startPos) else error "Маркеры данных не найдены в файле 1",

        promoted = Table.PromoteHeaders(dataRange, [PromoteAllScalars=true]),
		filtered = Table.SelectRows(promoted, each [Наименование] <> null),
		selected = Table.SelectColumns(filtered, {"Наименование", "Ед. измерения"}, MissingField.UseNull),

		// Очистка текста
		cleaned = Table.TransformColumns(selected, {
			{"Наименование", each Text.Clean(Text.Trim(_)), type text},
			{"Ед. измерения", each Text.Clean(Text.Trim(_)), type text}
		}),
        renamed = Table.RenameColumns(cleaned,{{"Наименование", "name"}, {"Ед. измерения", "capacityUnit"}})
        in
            renamed,

	// Обработка источника - DEMO-PACKAGING-STOCK-B
	consumable_2 = let
        fileContent = GetFileContent(FolderContent, filePart_2),
        workbook = Excel.Workbook(fileContent),
        rawData = GetSheetData(workbook, sheetName_2),

		startPos = List.PositionOf(rawData[Column1], headerMarker_2),
		validStart = startPos >= 0,
		dataRange = if validStart then Table.Skip(rawData, startPos + rowsToSkip_2) else error "Маркер данных не найден в файле 2",

		promoted = Table.PromoteHeaders(dataRange, [PromoteAllScalars=true]),
		filtered = Table.SelectRows(promoted, each [Column1] <> null),
		renamed = Table.RenameColumns(filtered, {{"Column2", "consumable"}, {"Column3", "capacityUnit"}}),

		// Динамический выбор столбцов
		allCols = Table.ColumnNames(renamed),
		keepCols = List.Select(allCols, each not Text.Contains(_, "Column", Comparer.OrdinalIgnoreCase)
											and not Text.Contains(_, "декада", Comparer.OrdinalIgnoreCase)),
		// Добавление ключевых столбцов
        finalCols = List.Combine({{"consumable", "capacityUnit"}, List.Difference(keepCols, {"consumable", "capacityUnit"})}),
		selected = Table.SelectColumns(renamed, finalCols, MissingField.UseNull),

		unpivoted = Table.UnpivotOtherColumns(selected, {"consumable", "capacityUnit"}, "Атрибут", "consumedQuantity"),

		cleaned = Table.TransformColumns(unpivoted, {
            {"consumable", each Text.Clean(Text.Trim(_)), type text},
            {"capacityUnit", each Text.Clean(Text.Trim(_)), type text}
        }),
		finalStep = Table.SelectColumns(cleaned, {"consumable", "capacityUnit"}),
        distinct = Table.Distinct(finalStep),
        result = Table.RenameColumns(distinct, {{"consumable", "name"}})
        in
            result,
	// Объединение запросов
	combined = Table.Combine({consumable_1, consumable_2}),
	typed = Table.TransformColumnTypes(combined, {{"name", type text}, {"capacityUnit", type text}}),
	finalTable = Table.Distinct(typed, {"name"})
in
     finalTable
