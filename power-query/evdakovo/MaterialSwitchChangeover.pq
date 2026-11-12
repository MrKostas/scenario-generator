let
    filteredFiles = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "матрица")),

    //функция обработки файла
    processFile = (currentFile) => let
        connectFile = currentFile[Content],
        dataTable = Excel.Workbook(connectFile),
        workSheets = Table.SelectRows(dataTable, each Text.Contains([Name], "Переход_Евдаково") and not Text.Contains([Name], "Print_Area")),

        // Функция обработки одного листа
        processSheet = (currentSheet) => let
            connectSheet = currentSheet[Data],

            lineValue = currentSheet[Column1]{List.PositionOf(List.Transform(currentSheet[Column1], each Text.Contains(_, "Матрица переходов")), true)},
            trimmedText = Text.Trim(Text.AfterDelimiter(lineValue, "участке")),
            finalLineValue = Text.Combine({"ЦПСЖ ", Text.BeforeDelimiter(trimmedText, " ", 1)}),

            numberOfStartRows = List.PositionOf(List.Transform(currentSheet[Column1], each Text.Contains(_, "с…на…")), true),
            numberOfEndRows = List.PositionOf(List.Transform(currentSheet[Column1], each Text.Contains(_, "Промывка")), true),

            filterRows = Table.FirstN(
                    currentSheet,
                    numberOfEndRows),
            filterRows2 = Table.Skip(filterRows, numberOfStartRows),
            promoteHeaders = Table.PromoteHeaders(filterRows2, [PromoteAllScalars=true]),

            selectColumns = Table.SelectColumns(
                promoteHeaders,
                List.Select(
                    Table.ColumnNames(promoteHeaders),
                    each not Text.Contains(_, "Column", Comparer.OrdinalIgnoreCase)
                )
            ),

            fillDownValues = Table.FillDown(selectColumns, Table.ColumnNames(selectColumns)),
            unpivotTable = Table.UnpivotOtherColumns(fillDownValues, {"с…на…"}, "Атрибут", "cleaningMaterial"),

            splitColumn = Table.ExpandListColumn(Table.TransformColumns(unpivotTable, {{"Атрибут", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}), "Атрибут"),
            renameColumn = Table.RenameColumns(splitColumn, {{"Атрибут", "nextMaterial"}}),

            selectMainColumns = Table.SelectColumns(renameColumn, {"с…на…", "cleaningMaterial", "nextMaterial"}),

            splitColumn2 = Table.ExpandListColumn(Table.TransformColumns(selectMainColumns, {{"с…на…", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}), "с…на…"),
            renameColumn2 = Table.RenameColumns(splitColumn2, {{"с…на…", "previousMaterial"}}),

            selectMainColumns2 = Table.SelectColumns(renameColumn2, {"previousMaterial", "nextMaterial", "cleaningMaterial"}),

            cleanText = Table.TransformColumns(selectMainColumns2,{{"previousMaterial", Text.Trim, type text}, {"nextMaterial", Text.Trim, type text}, {"cleaningMaterial", Text.Trim, type text}}),

            addLineColumn = Table.AddColumn(cleanText, "line", each finalLineValue, type text),
            addDurationColumn = Table.AddColumn(addLineColumn, "duration", each 1, Int64.Type),

            excludeNonUsefulRows = Table.SelectRows(addDurationColumn, each ([cleaningMaterial] <> "без промывки")),
            replaceValue = Table.ReplaceValue(excludeNonUsefulRows,"промывка ","",Replacer.ReplaceText,{"cleaningMaterial"})
        in
            replaceValue,
        processedSheets = Table.AddColumn(workSheets, "Processed", each processSheet([Data])),
        combinedResult = Table.Combine(processedSheets[Processed])
    in combinedResult,

    processedFiles = Table.AddColumn(filteredFiles, "ProcessedData", each processFile(_)),
    finalResult = Table.ExpandTableColumn(processedFiles, "ProcessedData", Table.ColumnNames(processedFiles{0}[ProcessedData])),
    selectFinalColumns = Table.SelectColumns(finalResult, {"duration", "line", "previousMaterial", "nextMaterial", "cleaningMaterial"}),
    finalFilterRows = Table.SelectRows(selectFinalColumns, each ([cleaningMaterial] <> "не планировать производство"))
in
    finalFilterRows
