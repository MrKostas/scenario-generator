let
    connect = Table.Distinct(InputRecord[[altInput], [processingStep]]),
    addAllColumns = Table.AddColumn(connect, "Custom", each [
        altInput = [altInput],
        beginDate = null,
        endDate = null,
        processingStep = [processingStp]
    ], type record),

    // Распаковка записи
    expandRecord = Table.ExpandRecordColumn(addAllColumns, "Custom",
    {"beginDate", "endDate"}
    ),

    // Финальная типизация
    convertTypes = Table.TransformColumnTypes(expandRecord, {
        {"altInput", type text},
        {"beginDate", type date},
        {"endDate", type date},
        {"processingStep", type text}
    }),
    renameAltColumn = Table.RenameColumns(
        convertTypes,
        {"altInput", "name"}
    ),
    addPriorityColumn = Table.AddColumn(renameAltColumn, "priority", each if Text.Contains([name], "Рецептура_1") then true else false),
    finalTable = Table.ReorderColumns(addPriorityColumn,{"name", "beginDate", "endDate", "priority", "processingStep"})
in
    finalTable
