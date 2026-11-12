let
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "переналадки")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Item="modeChange",Kind="Sheet"]}[Data],
    finalData = Table.PromoteHeaders(rawData, [PromoteAllScalars=true])
in
    finalData
