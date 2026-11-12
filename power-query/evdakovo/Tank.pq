let
    // connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "параметры мбх")){0}[Content],
    // dataTable = Excel.Workbook(connect),
    // filteredSheets = Table.SelectRows(dataTable, each [Kind] = "Sheet"),
    // sheetCount = Table.RowCount(filteredSheets),
    // lastSheetName = filteredSheets{sheetCount - 1}[Item],
    // rawData = dataTable{[Name = lastSheetName]}[Data],
    // numberOfSkipRows = List.PositionOf(rawData[Column1], "№ ёмкости"),
    // skipRowsPromoteHeaders = Table.PromoteHeaders(Table.Skip(rawData, numberOfSkipRows)),
    // selectColumns = Table.SelectColumns(skipRowsPromoteHeaders,{"№ ёмкости", "Кол-во заполнения при плотности 0,898 г/смᶟ (NAOH плотности 1,45 г/смᶟ)"}),
    // filterRows = Table.SelectRows(selectColumns, each ([#"№ ёмкости"] <> null)),
    // replaceTankNames = Table.ReplaceValue(filterRows,"DEMO/"," ",Replacer.ReplaceText,{"№ ёмкости"}),
    // replaceTankNames2 = Table.TransformColumns(replaceTankNames, {{"№ ёмкости", each "Бак " & _, type text}}),
    // transformValuesToTons = Table.TransformColumns(replaceTankNames2, {{"Кол-во заполнения при плотности 0,898 г/смᶟ (NAOH плотности 1,45 г/смᶟ)", each _ / 1000, Int64.Type}}),
    // changeColumnsNames = Table.RenameColumns(transformValuesToTons,{{"№ ёмкости", "name"}, {"Кол-во заполнения при плотности 0,898 г/смᶟ (NAOH плотности 1,45 г/смᶟ)", "capacity"}}),
    // roundValues = Table.TransformColumns(changeColumnsNames,{{"capacity", each Number.Round(_, 0), Int64.Type}}),
    // fantomTanks = #table(
    //     type table [
    //         name = Text.Type,
    //         capacity = Int64.Type
    //     ],
    //     {
    //         {"MAS E6", 1000000},
    //         {"Бак Лецитин", 1000000},
    //         {"MAS E7", 1000000},
    //         {"Бак 257_0", 1000000},
    //         {"Бак 258_0", 1000000},
    //         {"Вода", 10000000},
    //         {"Дубликат 1_Бак F 2", 4000000},
    //         {"Дубликат 1_Бак F 2/1", 4000000},
    //         {"Дубликат 2_Бак F 2", 4000000},
    //         {"Дубликат 2_Бак F 2/1", 4000000}
    //     }
    // ),
    // addFantomTanks = Table.Combine({roundValues, fantomTanks})
    connect = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), "трубы")){0}[Content],
    dataTable = Excel.Workbook(connect),
    rawData = dataTable{[Item="баки",Kind="Sheet"]}[Data],
    finalData = Table.PromoteHeaders(rawData, [PromoteAllScalars=true])
in
    finalData
