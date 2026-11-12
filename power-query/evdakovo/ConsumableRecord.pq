let
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "demo-consumption")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Name="ConsumableRecord"]}[Data],
    finalData = Table.PromoteHeaders(rawData, [PromoteAllScalars=true])
in
    finalData
