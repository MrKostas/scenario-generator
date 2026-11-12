let
    // параметры, участвующие в запросе
    fileNamePart = "demo-arrivals",
    sheetName = "МТТ",
    startMarker = "(каменка)",
    endMarker = "(на алексеевку)",
    headerSkipCount = 1,

    // безопасный выбор файла
    sourceFiles = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), Text.Lower(fileNamePart))),
    fileContent = if Table.RowCount(sourceFiles) > 0 then sourceFiles{0}[Content] else error "Файл не найден",

    // безопасный выбор листа
    workbookSheets = Excel.Workbook(fileContent),
    targetSheet = Table.SelectRows(workbookSheets, each [Item] = sheetName and [Kind] = "Sheet"),
    rawData = if Table.RowCount(targetSheet) > 0 then targetSheet{0}[Data] else error "Лист не найден",

    // поиск границ данных
    column1List = rawData[Column1],
    startPos = List.PositionOf(List.Transform(column1List, each Text.Contains(Text.Lower(_), startMarker)), true) - 1,
    startPosValid = if startPos = -1 then error "Маркер начала не найден" else startPos,

    endPos = List.PositionOf(List.Transform(column1List, each Text.Contains(Text.Lower(_), endMarker)), true),
    endPosValid = if endPos = -1 then error "Маркер конца не найден" else endPos,

    // Проверка на корректность позиций
    validRange = startPos >= 0 and endPos >= 0 and endPos > startPos,
    dataRange = Table.Range(rawData, startPosValid, endPosValid - startPosValid),

    // Подъем заголовков и пропуск служебных строк
    promotedTable = Table.PromoteHeaders(dataRange, [PromoteAllScalars=true]),
    cleanedTable = Table.Skip(promotedTable, headerSkipCount),

    // Очистка и переименование
    distinctTable = Table.Distinct(cleanedTable, {"Column1"}),
    renamedTable = Table.RenameColumns(distinctTable, {{"Column1", "material"}}),

    // Динамический выбор колонок
    columnNames = Table.ColumnNames(renamedTable),
    dataColumns = List.Select(columnNames, each not Text.StartsWith(_, "Column")),
    selectColumns = Table.SelectColumns(renamedTable, dataColumns),

    // Анпивот с типизацией
    unpivotTable = Table.UnpivotOtherColumns(selectColumns, {"material"}, "productionDate", "mass"),
    typedTable = Table.TransformColumnTypes(unpivotTable, {
        {"material", type text},
        {"productionDate", type date},
        {"mass", type number}
    }),

    // Фильтр по месяцу
    currentMonthStart = Date.AddMonths(DateTime.Date(Date.StartOfMonth(DateTime.LocalNow())), -2),
    monthFiltered = Table.SelectRows(typedTable, each Date.StartOfMonth([productionDate]) = currentMonthStart),

    // Добавление дат
    dateTable = Table.AddColumn(monthFiltered, "beginDate", each Date.AddDays([productionDate], 2), type date),
    dateTable2 = Table.AddColumn(dateTable, "endDate", each Date.AddDays([productionDate], 3), type date),

    // Подготовка к соединению (переименование для join)
    renamedForJoin = Table.RenameColumns(dateTable2, {{"material", "materialKey"}}),

    // Получение зон приемки
    receivingAreas = try Table.SelectRows(Material, each [unlimitedSupplyArea] <> null) otherwise #table({"name", "unlimitedSupplyArea"}, {}),
    receivingAreasClean = Table.SelectColumns(receivingAreas, {"name", "unlimitedSupplyArea"}),

    // Соединение таблиц
    joinResult = Table.Join(renamedForJoin, "materialKey", receivingAreasClean, "name", JoinKind.Inner),

    // Финальная организация столбцов
    finalTable = Table.SelectColumns(joinResult, {"beginDate", "endDate", "mass", "productionDate", "materialKey", "unlimitedSupplyArea"}),
    resultTable = Table.RenameColumns(finalTable, {{"materialKey", "material"}})
in
    resultTable
