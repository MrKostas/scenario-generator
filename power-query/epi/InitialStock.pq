let
    // Параметры, исползуемые в запросе
    fileNamePart = "demo-stock-b",
    sheetName = "Остатки",
    headerMarker = "№ п/п",
    rowsToSkip = 1,
    tankCodePrefix = "DEMO/",

    // Безопасный выбор файла
    sourceFiles = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), Text.Lower(fileNamePart))),
    fileContent = if Table.RowCount(sourceFiles) > 0 then sourceFiles{0}[Content] else error "Файл не найден",

    // Безопасный выбор листа
    workbookSheets = Excel.Workbook(fileContent),
    targetSheet = Table.SelectRows(workbookSheets, each [Item] = sheetName and [Kind] = "Sheet"),
    rawData = if Table.RowCount(targetSheet) > 0 then targetSheet{0}[Data] else error "Лист не найден",

    // Поиск заголовка с проверкой
    headerPos = List.PositionOf(rawData[Column2], headerMarker),
    headerPosValid = if headerPos = -1 then error "Заголовок не найден" else headerPos,

    // Пропуск строк и подъем заголовков
    promotedTable = Table.PromoteHeaders(Table.Skip(rawData, headerPos), [PromoteAllScalars=true]),
    skipTotal = Table.Skip(promotedTable, rowsToSkip),

    // Конвертация массы с обработкой ошибок
    removeErrors = Table.RemoveRowsWithErrors(skipTotal, {"Масса, кг"}),
    convertMass = Table.TransformColumns(removeErrors, {
        {"Масса, кг", each try Number.From(_) otherwise null, type number}
    }),

    // Фильрация null и 0
    filteredRows = Table.SelectRows(convertMass, each [#"Масса, кг"] <> null and [#"Масса, кг"] <> 0),

    // FillDown для групп
    fillDownTable = Table.FillDown(filteredRows, {"№ п/п", "Группа емкостей"}),

    // Очистка и переименования колонок
    cleanTank = Table.TransformColumns(fillDownTable, {
        {"№ емкости", each Text.Trim(Text.Replace(_, tankCodePrefix, " ")), type text}
    }),
    addTanksPrefix = Table.TransformColumns(cleanTank, {
        {"№ емкости", each "Бак " & _, type text}
    }),
    renameColumns = Table.RenameColumns(addTanksPrefix, {
        {"Группа емкостей", "tanksFunctionality"},
        {"Наименование сырья/продукции", "material"},
        {"№ емкости", "tank"},
        {"Масса, кг", "mass"}
    }),

    // Выбор колонок + типизация
    selectColumns = Table.SelectColumns(renameColumns, {"tanksFunctionality", "material", "tank", "mass"}),
    typedTable = Table.TransformColumnTypes(selectColumns, {
        {"tanksFunctionality", type text},
        {"material", type text},
        {"tank", type text},
        {"mass", type number}
    }),

    // Конвертация в тонны + добавление даты
    convertToTons = Table.TransformColumns(typedTable, {
        {"mass", each Number.Round(_ / 1000, 5), type number}
    }),
    productionDate = DateTime.From(Date.AddMonths(DateTime.Date(Date.StartOfMonth(DateTime.LocalNow())), -1)) + #duration(0, 8, 0, 0),
    addDate = Table.AddColumn(convertToTons, "productionDate", each productionDate, type datetime),

    // Фантомные баки
    fantomTanks = #table(
        type table [
            tanksFunctionality = Text.Type,
            material = Text.Type,
            tank = Text.Type,
            mass = number,
            productionDate = date
        ],
        {
            {"Фантомные баки", "MASESTER E6000", "MAS E6", 1000000, productionDate},
            {"Фантомные баки", "MASESTER E7000", "MAS E7", 1000000, productionDate},
            {"Фантомные баки", "Вода", "Вода", 10000000, productionDate}
        }
    ),
    combined = Table.Combine({addDate, fantomTanks}),

    // Финальный выбор столбцов
    finalTable = Table.SelectColumns(combined, {"mass", "productionDate", "tank", "material"})
in
    finalTable
