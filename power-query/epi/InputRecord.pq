let
    // Константы запроса
    RecipeHeaderMarker = "N/N",

    // Выбор файла с проверкой на существование
    sourceFiles = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "рецептуры")),
    fileContent = if Table.RowCount(sourceFiles) > 0 then sourceFiles{0}[Content] else error "Файл с рецептурами не найден",
    receiptTable = Table.PromoteHeaders(Excel.Workbook(FolderContent{List.PositionOf(FolderContent[Name], "DEMO-Рецепты без Дезо и ПереЭт.xlsx")}[Content])[Data]{0}, [PromoteAllScalars = true]),

    // Выбор последнего листа в книге
    workbookSheets = Excel.Workbook(fileContent),
    filteredSheets = Table.SelectRows(workbookSheets, each [Kind] = "Sheet"),
    targetSheet = if Table.RowCount(filteredSheets) > 0 then Table.Last(filteredSheets) else error "Листы не найдены",
    rawData = targetSheet[Data],

    // Поиск заголовка "N/N" и пропуск строк
    // Защита от случая, если заголовок не найден
    headerPos = List.PositionOf(rawData[Column1], RecipeHeaderMarker),
    rowsToSkip = if headerPos = -1 then 0 else headerPos,
    promotedTable = Table.PromoteHeaders(Table.Skip(rawData, rowsToSkip), [PromoteAllScalars=true]),

    selectColumns = Table.SelectColumns(promotedTable, {"НАИМЕНОВАНИЕ ПРОДУКЦИИ", "РЕЦЕПТ", "РЕЦЕПТУРЫ ПЕР.ЖИРОВ"}),
    filteredRows = Table.SelectRows(selectColumns, each not List.Contains(List.InsertRange(MaterialGroup[name], 0, {"Итого"}), [#"НАИМЕНОВАНИЕ ПРОДУКЦИИ"])),

    //Преобразования для пержиров
    pereTable = Table.SelectRows(filteredRows[[#"РЕЦЕПТУРЫ ПЕР.ЖИРОВ"]], each [#"РЕЦЕПТУРЫ ПЕР.ЖИРОВ"] <> null and [#"РЕЦЕПТУРЫ ПЕР.ЖИРОВ"] <> ""),
    pereExtractValue = Table.AddColumn(pereTable, "НАИМЕНОВАНИЕ ПРОДУКЦИИ", each Text.BetweenDelimiters([#"РЕЦЕПТУРЫ ПЕР.ЖИРОВ"], "%", " ")),
    pereSplitReceipt = Table.ExpandListColumn(Table.TransformColumns(pereExtractValue, {{"РЕЦЕПТУРЫ ПЕР.ЖИРОВ", Splitter.SplitTextByDelimiter(" / ", QuoteStyle.Csv), let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}), "РЕЦЕПТУРЫ ПЕР.ЖИРОВ"),
    pereExtractReceipt = Table.TransformColumns(pereSplitReceipt, {{"РЕЦЕПТУРЫ ПЕР.ЖИРОВ", each Text.BetweenDelimiters(_, "(", ")"), type text}}),
    pereDistinct = Table.Distinct(pereExtractReceipt, {"РЕЦЕПТУРЫ ПЕР.ЖИРОВ", "НАИМЕНОВАНИЕ ПРОДУКЦИИ"}),
    pereRenameColumn = Table.RenameColumns(pereDistinct,{{"РЕЦЕПТУРЫ ПЕР.ЖИРОВ", "РЕЦЕПТ"}}),
    pereAddProcess = Table.AddColumn(pereRenameColumn, "ПРОЦЕСС", each "ПереЭт"),

    // Преобразования для дезо
    dezoTable = filteredRows[[#"НАИМЕНОВАНИЕ ПРОДУКЦИИ"], [#"РЕЦЕПТ"]],
    dezoAddProcess = Table.AddColumn(dezoTable, "ПРОЦЕСС", each "Дезо"),

    // Объединение таблиц
    combineTable = Table.Combine({dezoAddProcess, pereAddProcess, receiptTable}, {"НАИМЕНОВАНИЕ ПРОДУКЦИИ", "РЕЦЕПТ", "ПРОЦЕСС"}),

    cleanedProduct = Table.TransformColumns(combineTable, {
        {"НАИМЕНОВАНИЕ ПРОДУКЦИИ", each Text.Clean(Text.Trim(_)), type text}
    }),

    // Объединение столбцов Процесс + Продукция
    combinedStep = Table.CombineColumns(cleanedProduct, {"ПРОЦЕСС", "НАИМЕНОВАНИЕ ПРОДУКЦИИ"},
        Combiner.CombineTextByDelimiter("_", QuoteStyle.None), "processingStep"),

    // Создание альтернативных рецептур
    splitAltInputs = Table.ExpandListColumn(Table.TransformColumns(combinedStep, {{"РЕЦЕПТ", Splitter.SplitTextByDelimiter(" / ", QuoteStyle.Csv), let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}), "РЕЦЕПТ"),
    addAltInputsNumbers = Table.Group(splitAltInputs, {"processingStep"}, {{"Data", each Table.AddIndexColumn(_, "Index", 1, 1, Int64.Type), type table}}),
    ungroupTable = Table.ExpandTableColumn(addAltInputsNumbers, "Data", {"РЕЦЕПТ", "Index"}, {"РЕЦЕПТ", "Index"}),
    addIndexColumn = Table.TransformColumns(ungroupTable, {{"Index", each "_Рецептура_" & Text.From(_)}}),
    addAltInputsColumn = Table.AddColumn(addIndexColumn, "altInput", each [processingStep] & [Index]),

    // Разбиение рецепта на список
    addedList = Table.AddColumn(addAltInputsColumn, "TempList",
        each List.Select(Text.Split([РЕЦЕПТ], " + "), each Text.Length(_) > 0),
        type list
    ),
    expandedList = Table.ExpandListColumn(addedList, "TempList"),

    // Разделение по "%" и переименование
    splitPercent = Table.SplitColumn(expandedList, "TempList",
        Splitter.SplitTextByDelimiter("%", QuoteStyle.Csv),
        {"consumedFraction", "material"}
    ),

    // Финальная типизация и расчет (Деление на 100 + приведение к числу)
    transformedValues = Table.TransformColumns(splitPercent, {
        {"consumedFraction", each Number.Round(Number.From(_) / 100, 3), type number},
        {"material", each Text.Trim(_), type text},
        {"processingStep", each Text.Trim(_), type text}
    }),
    reorderColumns = Table.SelectColumns(transformedValues,{"consumedFraction", "material", "altInput", "processingStep"})
in
    reorderColumns
