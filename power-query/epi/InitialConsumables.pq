let
    initialConsumables_1 = let
        connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "demo-consumable-balance")){0}[Content],
        dataTable = Excel.Workbook(connect),
        rawData = dataTable{[Name="DEMO-CONSUMABLE-BALANCE"]}[Data],
        numberOfStartRows = List.PositionOf(rawData[Column2], "№ п/п"),
        numberOfEndRows = List.PositionOf(List.Transform(rawData[Column3], each Text.Contains(_, "Исполнитель:")), true),
        skipRows = Table.Skip(rawData, numberOfStartRows),
        skipRows2 = Table.FirstN(skipRows, numberOfEndRows),
        promoteHeaders = Table.PromoteHeaders(skipRows2, [PromoteAllScalars=true]),
        renameConsumableColumn = Table.RenameColumns(promoteHeaders,{{"Column3", "consumables"}}),
        columnsList = Table.ColumnNames(renameConsumableColumn),
        indicesToKeep = {
            List.PositionOf(columnsList, "consumables"),
            List.PositionOf(columnsList, "consumables") + 1,
            List.PositionOf(columnsList, "ИТОГО") + 1},
        filtrationOfColumns = List.Transform(indicesToKeep, each columnsList{_}),
        chooseColumns = Table.SelectColumns(renameConsumableColumn, filtrationOfColumns),
        renameColumns = Table.RenameColumns(
            chooseColumns,
            {
                {Table.ColumnNames(chooseColumns){1}, "attribute"},
                {Table.ColumnNames(chooseColumns){2}, "initialStocks"}
            }
        ),
        fillDownConsumables = Table.FillDown(renameColumns, {"consumables"}),
        filterRows = Table.SelectRows(fillDownConsumables, each ([attribute] = "Ост.факт.по справке")),
        addDateColumn = Table.AddColumn(filterRows, "beginDate", each DateTime.From(Date.StartOfMonth(Date.AddMonths(DateTime.Date(DateTime.LocalNow()), -1))) + #duration(0, 8, 0, 0), type datetime),
        selectColumns = Table.SelectColumns(addDateColumn, {"consumables", "initialStocks", "beginDate"}),
        finalTable = Table.TransformColumnTypes(selectColumns,{{"initialStocks", type number}}),
        fillNullValues = Table.ReplaceValue(finalTable,null,0,Replacer.ReplaceValue,{"initialStocks"})
        in
            fillNullValues,
    initialConsumables_2 = let
        connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "demo-consumable-balance(1)")){0}[Content],
        dataTable = Excel.Workbook(connect),
        rawData = dataTable{[Name="DEMO-PERFORMANCE"]}[Data],
        numberOfStartRows = List.PositionOf(rawData[Column2], "№ п/п"),
        numberOfEndRows = List.PositionOf(List.Transform(rawData[Column3], each Text.Contains(_, "Исполнитель:")), true),
        skipRows = Table.Skip(rawData, numberOfStartRows),
        skipRows2 = Table.FirstN(skipRows, numberOfEndRows),
        promoteHeaders = Table.PromoteHeaders(skipRows2, [PromoteAllScalars=true]),
        renameConsumableColumn = Table.RenameColumns(promoteHeaders,{{"Наименование", "consumables"}}),
        columnsList = Table.ColumnNames(renameConsumableColumn),
        indicesToKeep = {
            List.PositionOf(columnsList, "consumables"),
            List.PositionOf(columnsList, "consumables") + 1,
            List.PositionOf(columnsList, "ИТОГО") + 1},
        filtrationOfColumns = List.Transform(indicesToKeep, each columnsList{_}),
        chooseColumns = Table.SelectColumns(renameConsumableColumn, filtrationOfColumns),
        renameColumns = Table.RenameColumns(
            chooseColumns,
            {
                {Table.ColumnNames(chooseColumns){1}, "attribute"},
                {Table.ColumnNames(chooseColumns){2}, "initialStocks"}
            }
        ),
        fillDownConsumables = Table.FillDown(renameColumns, {"consumables"}),
        filterRows = Table.SelectRows(fillDownConsumables, each ([attribute] = "Ост.факт.по справке")),
        addDateColumn = Table.AddColumn(filterRows, "beginDate", each DateTime.From(Date.StartOfMonth(Date.AddMonths(DateTime.Date(DateTime.LocalNow()), -1))) + #duration(0, 8, 0, 0), type datetime),
        selectColumns = Table.SelectColumns(addDateColumn, {"consumables", "initialStocks", "beginDate"}),
        finalTable = Table.TransformColumnTypes(selectColumns,{{"initialStocks", type number}}),
        fillNullValues = Table.ReplaceValue(finalTable,null,0,Replacer.ReplaceValue,{"initialStocks"})
        in
            fillNullValues,
    finalTable = Table.Combine({initialConsumables_1, initialConsumables_2}),
    roundValues = Table.TransformColumns(finalTable,{{"initialStocks", each Number.Round(_, 5), type number}})
in
    roundValues
