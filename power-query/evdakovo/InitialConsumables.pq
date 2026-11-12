let
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "demo-packaging-stock-a")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Name="InitialConsumables"]}[Data],
    finalData = Table.PromoteHeaders(rawData, [PromoteAllScalars=true])
in
    finalData
