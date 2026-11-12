let
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "лаборатория")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Item="LabStudyMode",Kind="Sheet"]}[Data],
    finalData = Table.PromoteHeaders(rawData, [PromoteAllScalars=true])
in
    finalData
