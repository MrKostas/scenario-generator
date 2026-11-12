let
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "demo-performance")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Item="SimpleProcessingStep",Kind="Sheet"]}[Data],
    finalData = Table.PromoteHeaders(rawData, [PromoteAllScalars=true])
in
    finalData
