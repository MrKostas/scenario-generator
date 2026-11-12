let
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "demo-operations")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Item="PerformanceRecord",Kind="Sheet"]}[Data],
    finalData = Table.PromoteHeaders(rawData, [PromoteAllScalars=true])
in
    finalData
