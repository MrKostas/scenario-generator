let
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "demo-performance")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Item="ComplexStep",Kind="Sheet"]}[Data],
    finalData = Table.PromoteHeaders(rawData, [PromoteAllScalars=true]),
    convertColumnTypes = Table.TransformColumnTypes(finalData,{{"name", type text}, {"producedFraction1", type number}, {"producedFraction2", type number}, {"producedMaterial1", type text}, {"producedMaterial2", type text}})
in
    convertColumnTypes
