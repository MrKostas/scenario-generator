let
    // connect = Line,
    // selectRows = Table.SelectRows(connect, each Text.Contains([name], "МСлив", Comparer.OrdinalIgnoreCase)),
    // ExtractBase = Table.AddColumn(selectRows, "БазоваяЧасть", each
    //     let
    //         parts = Text.Split([name], " "),
    //         lastPart = parts{List.Count(parts) - 1},
    //         isNumber = Text.StartsWith(lastPart, "№")
    //     in
    //         if isNumber then
    //             Text.Combine(List.RemoveLastN(parts, 1), " ")
    //         else
    //             [line]
    // ),

    // AddSpecialFlag = Table.AddColumn(ExtractBase, "IsSpecial", each
    //     Text.Contains([name], "МСлив 2_Авто №3", Comparer.OrdinalIgnoreCase) or
    //     Text.Contains([name], "МСлив 2_Авто №4", Comparer.OrdinalIgnoreCase)
    // ),

    // NormalRows = Table.SelectRows(AddSpecialFlag, each not [IsSpecial]),
    // SpecialRows = Table.SelectRows(AddSpecialFlag, each [IsSpecial]),

    // UniqueNormalBases = Table.Distinct(Table.SelectColumns(NormalRows, {"БазоваяЧасть"})),
    // AddNormalZoneNumber = Table.AddIndexColumn(UniqueNormalBases, "ЗонаНомер", 1, 1),
    // NormalZoneMapping = Table.AddColumn(AddNormalZoneNumber, "LineGroup", each "Зона приемки_" & Text.From([ЗонаНомер])),

    // MergedNormal = Table.NestedJoin(
    //     NormalRows,
    //     {"БазоваяЧасть"},
    //     NormalZoneMapping,
    //     {"БазоваяЧасть"},
    //     "ЗонаДанные",
    //     JoinKind.LeftOuter
    // ),
    // ExpandedNormal = Table.ExpandTableColumn(MergedNormal, "ЗонаДанные", {"LineGroup"}, {"LineGroup"}),

    // ExpandedSpecial = Table.ExpandListColumn(
    //     Table.AddColumn(SpecialRows, "LineGroup", each {"ЦФСЖ Авто №6", "ЦФСЖ Авто №7"}),
    //     "LineGroup"
    // ),

    // CombinedResult = Table.Combine({ExpandedNormal, ExpandedSpecial}),

    // Result = Table.SelectColumns(CombinedResult, {"LineGroup", "name"}),
    // renameColumns = Table.RenameColumns(Result,{{"name", "lines"}})
    connect = Table.FromRecords({[name = null,
                                    lines = null]}),
    convertTypes = Table.TransformColumnTypes(connect,{{"name", type text},
                                                        {"lines", type text}})
in
    convertTypes
