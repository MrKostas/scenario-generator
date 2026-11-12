let
    connect = TankGroup_tanks[TankGroup],
    transformToTable = Table.FromList(connect, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    convertColumnTypes = Table.TransformColumnTypes(
        Table.RenameColumns(transformToTable,{{"Column1", "name"}}),
        {{"name", type text}}),
    dropDuplicates = Table.Distinct(convertColumnTypes)
in
    dropDuplicates
