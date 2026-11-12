let
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "demo-performance")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Item="PerformanceRecord",Kind="Sheet"]}[Data],
    finalData = Table.PromoteHeaders(rawData, [PromoteAllScalars=true]),
    roundValues = Table.TransformColumns(finalData,
                                        {{"summerPerformance", each Number.Round(_, 5), type number},
                                        {"winterPerformance", each Number.Round(_, 5), type number},
                                        {"minBatchSize", each Number.Round(_, 0), Int64.Type},
                                        {"batchMultiple", each Number.Round(_, 0), Int64.Type},
                                        {"subStep", each Number.Round(_, 0), Int64.Type}})
in
    roundValues
