let
    // параметры, участвующие в запросе
    fileNamePart = "demo-arrivals",
    sheetName = "МТТ",
    startMarker = "(алексеевка)",
    endMarker = "(каменка)",
    headerSkipCount = 3,

    // безопасный выбор файла
    sourceFiles = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), Text.Lower(fileNamePart))),
    fileContent = if Table.RowCount(sourceFiles) > 0 then sourceFiles{0}[Content] else error "Файл не найден",

    // безопасный выбор листа
    workbookSheets = Excel.Workbook(fileContent),
    targetSheet = Table.SelectRows(workbookSheets, each [Item] = sheetName and [Kind] = "Sheet"),
    rawData = if Table.RowCount(targetSheet) > 0 then targetSheet{0}[Data] else error "Лист не найден",

    // поиск границ данных
    column1List = rawData[Column1],
    startPos = List.PositionOf(List.Transform(column1List, each Text.Contains(Text.Lower(_), "(алексеевка)")), true) - 1,
    startPosValid = if startPos = -1 then error "Маркер начала не найден" else startPos,

    endPos = List.PositionOf(List.Transform(column1List, each Text.Contains(Text.Lower(_), "(каменка)")), true),
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
    currentMonthStart = Date.AddMonths(DateTime.Date(Date.StartOfMonth(DateTime.LocalNow())), -1),
    monthFiltered = Table.SelectRows(typedTable, each Date.StartOfMonth([productionDate]) = currentMonthStart),

    // Добавление дат
    dateTable = Table.AddColumn(monthFiltered, "beginDate", each DateTime.From(Date.AddDays([productionDate], 2)) + #duration(0, 8, 0, 0), type datetime),
    dateTable2 = Table.AddColumn(dateTable, "endDate", each DateTime.From(Date.AddDays([productionDate], 3)) + #duration(0, 8, 0, 0), type datetime),

    // Подготовка к соединению (переименование для join)
    renamedForJoin = Table.RenameColumns(dateTable2, {{"material", "materialKey"}}),

    // Получение зон приемки
    receivingAreas = try Table.SelectRows(Material, each [unlimitedSupplyArea] <> null) otherwise #table({"name", "receivingArea"}, {}),
    receivingAreasClean = Table.SelectColumns(Table.RenameColumns(receivingAreas, {{"unlimitedSupplyArea", "receivingArea"}}), {"name", "receivingArea"}),

    // Соединение таблиц
    joinResult = Table.Join(renamedForJoin, "materialKey", receivingAreasClean, "name", JoinKind.Inner),

    // Финальная организация столбцов
    finalTable = Table.SelectColumns(joinResult, {"beginDate", "endDate", "mass", "productionDate", "materialKey", "receivingArea"}),
    resultTable = Table.RenameColumns(finalTable, {{"materialKey", "material"}}),
    roundValues = Table.TransformColumns(resultTable,{{"mass", each Number.Round(_, 5), type number}})
in
    roundValues
