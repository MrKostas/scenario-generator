let
    // Параметры, исползуемые в запросе
    fileNamePart = "demo-stock-a",
    sheetName = "тбх",
    headerMarker = "№",
    rowsToSkip = 1,
    // tankCodePrefix = "DEMO/",

    // Безопасный выбор файла
    sourceFiles = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), Text.Lower(fileNamePart))),
    fileContent = if Table.RowCount(sourceFiles) > 0 then sourceFiles{0}[Content] else error "Файл не найден",

    // Безопасный выбор листа
    workbookSheets = Excel.Workbook(fileContent),
    targetSheet = Table.SelectRows(workbookSheets, each [Name] = sheetName),
    rawData = if Table.RowCount(targetSheet) > 0 then targetSheet{0}[Data] else error "Лист не найден",

    // Поиск заголовка с проверкой
    headerPos = List.PositionOf(rawData[Column1], headerMarker),
    headerPosValid = if headerPos = -1 then error "Заголовок не найден" else headerPos,

    // Пропуск строк и подъем заголовков
    promotedTable = Table.PromoteHeaders(Table.Skip(rawData, headerPos), [PromoteAllScalars=true]),
    skipTotal = Table.Skip(promotedTable, rowsToSkip),

    // Конвертация массы с обработкой ошибок
    removeErrors = Table.RemoveRowsWithErrors(skipTotal, {"остаток сырья, (кг)"}),
    convertMass = Table.TransformColumns(removeErrors, {
        {"остаток сырья, (кг)", each try Number.From(_) otherwise null, type number}
    }),

    // Фильрация null и 0
    filteredRows = Table.SelectRows(convertMass, each [#"остаток сырья, (кг)"] <> null and [#"остаток сырья, (кг)"] <> 0),

    // Очистка и переименования колонок
    addTanksPrefix = Table.TransformColumns(filteredRows, {
        {"№ емкости", each "Бак " & _, type text}
    }),
    renameColumns = Table.RenameColumns(addTanksPrefix, {
        {"Наименование сырья и  продукции", "material"},
        {"№ емкости", "tank"},
        {"остаток сырья, (кг)", "mass"}
    }),

    // Выбор колонок + типизация
    selectColumns = Table.SelectColumns(renameColumns, {"material", "tank", "mass"}),
    typedTable = Table.TransformColumnTypes(selectColumns, {
        {"material", type text},
        {"tank", type text},
        {"mass", type number}
    }),

    // Конвертация в тонны + добавление даты
    convertToTons = Table.TransformColumns(typedTable, {
        {"mass", each _ / 1000, type number}
    }),
    productionDate = Date.AddMonths(DateTime.Date(Date.StartOfMonth(DateTime.LocalNow())), -1),
    addDate = Table.AddColumn(convertToTons, "productionDate", each productionDate, type date),

    // Фантомные баки
    fantomTanks = #table(
        type table [
            material = Text.Type,
            tank = Text.Type,
            mass = number,
            productionDate = date
        ],
        {
            // {"MASESTER E6000", "MAS E6", 1000000, productionDate},
            // {"MASESTER E7000", "MAS E7", 1000000, productionDate},
            {"Вода", "Вода", 10000000, productionDate}
        }
    ),
    combined = Table.Combine({addDate, fantomTanks}),

    // Финальный выбор столбцов
    finalTable = Table.SelectColumns(combined, {"mass", "productionDate", "tank", "material"})
in
    finalTable
