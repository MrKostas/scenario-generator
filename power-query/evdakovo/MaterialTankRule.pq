let
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "хранения")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Item="хранениеПродуктов",Kind="Sheet"]}[Data],
    promoteHeaders = Table.PromoteHeaders(rawData, [PromoteAllScalars=true]),
    selectColumns = Table.SelectColumns(promoteHeaders,{"supportsDuplexMode", "supportsConcurrentFilling", "material", "tank"})
in
    selectColumns
