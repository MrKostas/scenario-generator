let
    connect = Table.FromRecords({[name = "scenario name",
                                    beginDate = Date.AddMonths(Date.StartOfMonth(DateTime.Date(DateTime.LocalNow())), -1),
                                    endDate = Date.AddMonths(Date.EndOfMonth(DateTime.Date(DateTime.LocalNow())), -1),
                                    firstSummerMonth = 5,
                                    firstWinterMonth = 10,
                                    pumpingRate = 25]}),
    convertTypes = Table.TransformColumnTypes(connect,{{"name", type text},
                                                        {"beginDate", type date},
                                                        {"endDate", type date},
                                                        {"firstSummerMonth", Int64.Type},
                                                        {"firstWinterMonth", Int64.Type},
                                                        {"pumpingRate", Int64.Type}})
in
    convertTypes
