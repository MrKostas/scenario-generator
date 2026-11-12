let
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "трубы")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Item="группыБаков",Kind="Sheet"]}[Data],
    promoteHeaders = Table.TransformColumnTypes(
        Table.PromoteHeaders(rawData),
        {{"TankGroup", type text}, {"tanks", type text}})
in
    promoteHeaders
