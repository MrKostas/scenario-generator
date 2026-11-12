let
    connect = Material[aggregateStateCategory],
    dropDuplicates = List.Distinct(connect),
    transformToTable = Table.FromList(dropDuplicates, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    excludeNullRows = Table.SelectRows(transformToTable, each ([Column1] <> null and [Column1] <> "")),
    renameColumns = Table.RenameColumns(excludeNullRows,{{"Column1", "name"}})
in
    renameColumns
