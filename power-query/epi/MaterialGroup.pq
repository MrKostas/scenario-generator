let
    otherCategories = Table.FromList(
        {"Кристаллизаторы", "Эффективные корма", "Эмульгаторы"},
        null,
        {"name"}
    ),
    connect = MaterialGroup_materials[MaterialGroup],
    transformToTable = Table.FromList(connect, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    convertColumnTypes = Table.TransformColumnTypes(
        Table.RenameColumns(transformToTable,{{"Column1", "name"}}),
        {{"name", type text}}),
    dropDuplicates = Table.Distinct(convertColumnTypes),
    finalTable = Table.Combine({dropDuplicates, otherCategories}, {"name"})
in
    finalTable
