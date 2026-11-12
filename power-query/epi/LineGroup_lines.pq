let
    connect = Line,
    selectRows = Table.SelectRows(connect, each Text.Contains([name], "МСлив", Comparer.OrdinalIgnoreCase)),
    // 3. Извлекаем "базовую часть" для группировки (удаляем номер в конце: №1, №2 и т.п.)
    ExtractBase = Table.AddColumn(selectRows, "БазоваяЧасть", each
        let
            parts = Text.Split([name], " "),
            lastPart = parts{List.Count(parts) - 1},
            isNumber = Text.StartsWith(lastPart, "№")
        in
            if isNumber then
                Text.Combine(List.RemoveLastN(parts, 1), " ")
            else
                [line]
    ),
    // 4. Добавляем флаг для особых случаев
    AddSpecialFlag = Table.AddColumn(ExtractBase, "IsSpecial", each
        Text.Contains([name], "МСлив 2_Авто №3", Comparer.OrdinalIgnoreCase) or
        Text.Contains([name], "МСлив 2_Авто №4", Comparer.OrdinalIgnoreCase)
    ),

    // 5. Разделяем таблицу на обычные и особые строки
    NormalRows = Table.SelectRows(AddSpecialFlag, each not [IsSpecial]),
    SpecialRows = Table.SelectRows(AddSpecialFlag, each [IsSpecial]),

    // 6. Обрабатываем обычные строки (как раньше)
    UniqueNormalBases = Table.Distinct(Table.SelectColumns(NormalRows, {"БазоваяЧасть"})),
    AddNormalZoneNumber = Table.AddIndexColumn(UniqueNormalBases, "ЗонаНомер", 1, 1),
    NormalZoneMapping = Table.AddColumn(AddNormalZoneNumber, "LineGroup", each "Зона приемки_" & Text.From([ЗонаНомер])),

    MergedNormal = Table.NestedJoin(
        NormalRows,
        {"БазоваяЧасть"},
        NormalZoneMapping,
        {"БазоваяЧасть"},
        "ЗонаДанные",
        JoinKind.LeftOuter
    ),
    ExpandedNormal = Table.ExpandTableColumn(MergedNormal, "ЗонаДанные", {"LineGroup"}, {"LineGroup"}),

    // 7. Обрабатываем особые строки: создаем две копии с разными LineGroup
    ExpandedSpecial = Table.ExpandListColumn(
        Table.AddColumn(SpecialRows, "LineGroup", each {"ЦФСЖ Авто №6", "ЦФСЖ Авто №7"}),
        "LineGroup"
    ),

    // 8. Объединяем результаты
    CombinedResult = Table.Combine({ExpandedNormal, ExpandedSpecial}),

    // 9. Формируем итоговую таблицу
    Result = Table.SelectColumns(CombinedResult, {"LineGroup", "name"}),
    renameColumns = Table.RenameColumns(Result,{{"name", "lines"}})
in
    renameColumns
