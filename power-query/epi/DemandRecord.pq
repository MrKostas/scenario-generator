let
    // Параметры, исползуемые в запросе
    fileNamePart = "demo-demand",
    sheetName = "итого фасовка",
    headerMarker = "Итого фасовка",
    rowsToSkip = 1,
    maxDelay = 120,
    priority = 1000,

    // Безопасный выбор файла
    sourceFiles = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), fileNamePart)),
    fileContent = if Table.RowCount(sourceFiles) > 0 then sourceFiles{0}[Content] else error "Файл не найден",

    // Безопасный выбор листа
    workbookSheets = Excel.Workbook(fileContent),
    targetSheet = Table.SelectRows(workbookSheets, each Text.Lower([Name]) = Text.Lower(sheetName)),
    rawData = if Table.RowCount(targetSheet) > 0 then targetSheet{0}[Data] else error "Лист не найден",

    // Поиск заголовка и подготовка данных
    headerPos = List.PositionOf(rawData[Column1], headerMarker),
    validHeader = headerPos >= 0,
    dataRange = if validHeader then Table.Skip(rawData, headerPos) else error "Маркер не найден",
    finalStructure = Table.PromoteHeaders(Table.Skip(dataRange, rowsToSkip), [PromoteAllScalars=true]),

    // Выбор и переименование колонок
    selectColumns = Table.SelectColumns(finalStructure, {"Column1", "Итого за месяц"}, MissingField.UseNull),
    renamedTable = Table.TransformColumnTypes(Table.RenameColumns(selectColumns, {{"Column1", "material"}, {"Итого за месяц", "mass"}}), {{"mass", type number}}),

    // Фильтрация и очистка
    dropNull = Table.SelectRows(renamedTable, each ([mass] <> null and [mass] <> 0 and [material] <> null)),
    cleanMaterials = Table.TransformColumns(dropNull, {
        {"material", each Text.Clean(Text.Trim(_)), type text},
        {"mass", each Number.From(_), type number}
    }),
    filterRows = Table.SelectRows(cleanMaterials, each not List.Contains(List.InsertRange(MaterialGroup[name], 0, {"Итого", "Исполнитель: Демо"}), [material])),
    addIndexCol = Table.AddIndexColumn(filterRows, "index", 1),

    // Добавление всех колонок одним шагом
    currentDate = DateTime.Date(DateTime.LocalNow()),
    monthStart = DateTime.From(Date.AddMonths(Date.StartOfMonth(currentDate), -1)) + #duration(0, 8, 0, 0),
    monthEnd = DateTime.From(Date.AddDays(Date.AddMonths(Date.EndOfMonth(currentDate), -1), 2)) + #duration(0, 8, 0, 0),

    addAllColumns = Table.AddColumn(addIndexCol, "Custom", each [
        index = [index],
        name = "Заказ №" & Text.From([index]),
        enabled= true,
        material = [material],
        beginDate = monthStart,
        endDate = monthEnd,
        maxDelay = maxDelay,
        priority = priority,
        mass = [mass]
    ], type record),

    // Распаковка записи и сортировка
    expandRecord = Table.ExpandRecordColumn(addAllColumns, "Custom",
    {"name", "enabled", "beginDate", "endDate", "maxDelay", "priority"}
    ),
    sortedTable = Table.Sort(expandRecord, {{"mass", Order.Descending}, {"material", Order.Ascending}}),

    // Финальная типизация
    finalTable = Table.TransformColumnTypes(sortedTable, {
        {"name", type text},
        {"enabled", type logical},
        {"mass", type number},
        {"beginDate", type datetime},
        {"endDate", type datetime},
        {"maxDelay", Int64.Type},
        {"priority", Int64.Type},
        {"material", type text}
    }),
    reorderColumns = Table.SelectColumns(
        finalTable,
        {"name", "enabled", "mass", "beginDate", "endDate", "maxDelay", "priority", "material"}
    )
in
    reorderColumns
