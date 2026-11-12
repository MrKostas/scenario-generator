let
    // connect = List.Distinct(LineGroup_lines[LineGroup]),
    // transformToTable = Table.FromList(connect, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    // renameColumns = Table.RenameColumns(transformToTable,{{"Column1", "name"}})
    connect = Table.FromRecords({[name = null]}),
    convertTypes = Table.TransformColumnTypes(connect,{{"name", type text}})
in
    convertTypes
