let
    connect = Pipe,
    sourceValues = Table.Column(connect, "source"),
    destValues = Table.Column(connect, "dest"),
    CombinedList = List.Combine({sourceValues, destValues}),
    UniqueList = List.Distinct(CombinedList),
    Result = Table.FromList(UniqueList, Splitter.SplitByNothing(), {"line"}, null, ExtraValues.Error),
    selectRows = Table.SelectRows(Result, each not Text.Contains([line], "Бак") and not Text.Contains([line], "MAS", Comparer.OrdinalIgnoreCase) and Text.Length([line]) > 2 and [line] <> "Вода"),
    ValuesList = {
        "ЦВП Электролизный",
        "ЦВП Метанольный"
    },
    transformTable = Table.FromList(ValuesList, Splitter.SplitByNothing(), {"line"}, null, ExtraValues.Error),
    combinedTables = Table.Combine({selectRows, transformTable}),
    addColumnBusy = Table.AddColumn(combinedTables, "mustBeBusy", each if Text.Contains([line], "Дезо") and not Text.Contains([line], "Смеситель") then true else false),
    addColumnSection = Table.AddColumn(addColumnBusy, "lineSection",
    each if Text.BeforeDelimiter([line], " ") = "Зона"
          or Text.BeforeDelimiter([line], " ") = "Смеситель"
    then "Другое"
    else Text.BeforeDelimiter([line], " ")
),
    finalTable = Table.RenameColumns(addColumnSection, {{"line", "name"}})
in
    finalTable
