let
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "demo-packaging-stock-b")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Item="Остатки ",Kind="Sheet"]}[Data],
    numberOfSkipRows = List.PositionOf(rawData[Column1], "№ п/п"),
    skipRows = Table.PromoteHeaders(Table.Skip(rawData, numberOfSkipRows + 1)),
    filterNullRows = Table.SelectRows(skipRows, each ([Column1] <> null)),
    renameColumns = Table.RenameColumns(filterNullRows,{{"Column2", "consumable"}, {"Column3", "capacityUnit"}}),
    filteredColumnNames = Table.SelectColumns(
        renameColumns,
        List.Select(
            Table.ColumnNames(renameColumns),
            each not Text.Contains(_, "Column", Comparer.OrdinalIgnoreCase) and not Text.Contains(_, "декада", Comparer.OrdinalIgnoreCase)
        )
    ),
    unpivotOtherColumns = Table.UnpivotOtherColumns(filteredColumnNames, {"capacityUnit", "consumable"}, "Атрибут", "consumedQuantity"),
    addStepColumn = Table.AddColumn(unpivotOtherColumns, "step", each if Text.Contains([Атрибут], "налив") then "Слив_" & [Атрибут] else "Фас_" & [Атрибут], type text),
    addProdQuantity = Table.AddColumn(addStepColumn, "producedQuantity", each 1, Int64.Type),
    predfinalTable = Table.SelectColumns(addProdQuantity, {"consumedQuantity", "producedQuantity", "step", "consumable"}),
    linesValues = {"ЦФСЖ №1", "ЦФСЖ №2", "ЦФСЖ №3"},
    cols = Table.ToColumns(predfinalTable),
    newColumns = cols & {List.Repeat(linesValues, Number.RoundUp(Table.RowCount(predfinalTable) / 3))},
    finalTable = Table.SelectRows(Table.FromColumns(newColumns, Table.ColumnNames(predfinalTable) & {"line"}), each [consumedQuantity] <> null),
    reorderColumns = Table.ReorderColumns(finalTable,{"consumedQuantity", "producedQuantity", "line", "step", "consumable"}),
    roundValues = Table.TransformColumns(reorderColumns,{{"consumedQuantity", each Number.Round(_, 5), type number}})
in
    roundValues
