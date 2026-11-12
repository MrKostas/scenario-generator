let
    connect = List.Distinct(LineGroup_lines[LineGroup]),
    transformToTable = Table.FromList(connect, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    renameColumns = Table.RenameColumns(transformToTable,{{"Column1", "name"}})
in
    renameColumns
