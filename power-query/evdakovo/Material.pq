let
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "информация о продуктах")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Item="materials",Kind="Sheet"]}[Data],
    predFinalData = Table.PromoteHeaders(rawData, [PromoteAllScalars=true]),
    addPackagingColumn = Table.AddColumn(predFinalData, "Packaging", each if [technologyCategory] = "Отгруженные наливом" or [technologyCategory] = "Фасованные жиры" then true else false)
in
    addPackagingColumn
