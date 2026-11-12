let
  // Параметры запроса
  workbookName = "demo-maintenance-b",
  excludedEquipment = {"МБХ", "ЦГРМ", "ЦПВ №1", "ЦПВ №2", "ЦФСЖ участок №2 на 2026 год"},
  headerMarker = "Месяц",
  sheetNameMarker = "ремонта",
  referenceYear = Date.Year(DateTime.LocalNow()),

  // Безопасное подключение к файлу
  sourceFiles = Table.SelectRows(FolderContent, each Text.Contains(Text.Lower([Name]), workbookName)),
  fileContent = if Table.RowCount(sourceFiles) > 0 then sourceFiles{0}[Content] else error "Файл не найден",

  // Выбор листов
  workbookSheets = Excel.Workbook(fileContent),
  filteredSheets = Table.SelectRows(workbookSheets, each not Text.Contains([Name], "Print_Area") and Text.Lower([Name]) <> "инструкция"),

  // Обработка каждого листа
  processedSheets = Table.TransformColumns(
      filteredSheets,
      {"Data", each
        let
          currentTable = _,
          headerRowIdx = List.PositionOf(currentTable[Column1], headerMarker),
          isValidHeader = headerRowIdx >= 0,

          // Извлечение имени оборудования из имени листа или ячейки
          // Если заголовка нет, то пропускаем обработку листа
          sheetNameRaw = if isValidHeader then try (Text.AfterDelimiter(currentTable[Column2]{headerRowIdx}, sheetNameMarker)) otherwise null else error "Отсутствует нужное название заголовка",
          equipmentName = if sheetNameRaw <> null then Text.Trim(sheetNameRaw) else Text.Trim([Name]),

          // Подготовка таблицы данных
          skipRows = Table.Skip(currentTable, headerRowIdx + 1),
          promotedTable = Table.PromoteHeaders(skipRows, [PromoteAllScalars=true]),

          // Фильтр по месяцам
          monthList = {"январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"},
          filterRows = Table.SelectRows(promotedTable, each List.Contains(monthList, Text.Lower([Column1]))),

          // Переименование и очистка колонок
          renameMainColumn = Table.RenameColumns(filterRows, {{"Column1", "month"}}),
          // Оставляем только месяц и данные (исключаем служебные Column...)
          keepCols = List.Select(Table.ColumnNames(renameMainColumn), each _ <> "month" and not Text.StartsWith(_, "Column")),
          selectCols = Table.SelectColumns(renameMainColumn, List.Combine({{"month"}, keepCols})),

          // Pacpivot (преобразование в длинный формат)
          unpivotTable = Table.UnpivotOtherColumns(selectCols, {"month"}, "day", "description"),

          // Формирование даты
          addDateColumn = Table.AddColumn(unpivotTable, "date", each [day] & "." & [month] & "." & Text.From(Date.Year(DateTime.LocalNow()))),
          convertColumnTypes = Table.TransformColumnTypes(addDateColumn,{{"date", type date}}),
          addIndex= Table.AddIndexColumn(convertColumnTypes, "Index", 0, 1, Int64.Type),

          addPrevMonth = Table.AddColumn(addIndex, "PrevMonth", each try Date.Month(addIndex[date]{[Index]-1}) otherwise null, type text),
          addPrevDay = Table.AddColumn(addPrevMonth, "PrevDay", each try Date.Day(addIndex[date]{[Index]-1}) otherwise null, Int64.Type),
          addPrevDesc = Table.AddColumn(addPrevDay, "PrevDesc", each try addIndex[description]{[Index]-1} otherwise null, type text),

          //определяем флаг начала новой группы (1 - новая группа, 0 - продолжение)
          addFlag = Table.AddColumn(addPrevDesc, "IsNewGroup", each
              if [Index] = 0 then 1
              else if Date.Month([date]) <> [PrevMonth] then 1
              else if Date.Day([date]) <> [PrevDay] + 1 then 1
              else if [description] <> [PrevDesc] then 1
              else 0
          , Int64.Type),

          // формирование идентификатора (метки) как накопительную сумму флага
          flagList = addFlag[IsNewGroup],
          labelList = List.Accumulate(flagList, {}, (state, current) =>
              state & {(if List.Count(state) = 0 then current else List.Last(state) + current)}
          ),
          addLabel = Table.AddColumn(addFlag, "Label", each labelList{[Index]}, Int64.Type),

          // удаление вспомогательных столбцов
          removedHelpers = Table.RemoveColumns(addLabel, {"Index", "PrevMonth", "PrevDay", "PrevDesc", "IsNewGroup"}),

          groupTable = Table.Group(removedHelpers,
                            {"month", "description", "Label"},
                            {
                                {"beginDate", each List.Min([date]), type nullable date},
                                {"endDate", each List.Max([date]), type nullable date}
                            }),

          // добавление имени оборудования
          addEquipment = Table.AddColumn(groupTable, "equipment", each equipmentName, type text)
        in
          addEquipment
      }
  ),

  // Объединение всех листов
  combined = Table.Combine(processedSheets[Data]),

  // Динамическое разбиение оборудования (вместо жесткой фиксации колонок)
  // Дробление текста на список и его развертка
  splitEquipment = Table.AddColumn(combined, "EquipmentList", each Text.Split([equipment], ", ")),
  expandedEquipment = Table.ExpandListColumn(splitEquipment, "EquipmentList"),
  renamedEquipment = Table.RenameColumns(expandedEquipment, {{"EquipmentList", "sections"}}),
  trimmedEquipment = Table.TransformColumns(renamedEquipment, {{"equipment", Text.Trim, type text}}),

  // Фильтрация (список исключений + Дата)
  filterEquipment = Table.SelectRows(trimmedEquipment, each not List.Contains(excludedEquipment, [sections])),

  currentMonthStart = Date.AddMonths(Date.StartOfMonth(DateTime.Date(DateTime.LocalNow())), -1),
  filterDate = Table.SelectRows(filterEquipment, each Date.StartOfMonth([beginDate]) = currentMonthStart),
  selectColumns = Table.SelectColumns(filterDate, {"beginDate", "endDate", "description", "sections"}),
  finalTable = Table.RenameColumns(selectColumns, {{"sections", "equipment"}}),
  changeDateTypes = Table.TransformColumns(
    finalTable,
    {
        {"beginDate", each DateTime.From(_) + #duration(0, 8, 0, 0), type datetime},
        {"endDate", each DateTime.From(Date.AddDays(_, 1)) + #duration(0, 8, 0, 0), type datetime}
    }
  )
in
  changeDateTypes
