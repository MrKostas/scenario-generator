let
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "demo-consumables")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Item="Consumable",Kind="Sheet"]}[Data],
    finalData = Table.PromoteHeaders(rawData, [PromoteAllScalars=true])
in
    finalData
