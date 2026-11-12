let
    connect = Material[[endProductCategory], [name], [technologyCategory]],
    selectRows = Table.SelectRows(connect, each ([technologyCategory] = "Дезодорированные жиры" or [technologyCategory] = "Переэтерифицированные жиры")),
    renameColumns = Table.RenameColumns(selectRows, {{"endProductCategory", "MaterialGroup"}, {"name", "materials"}}),
    finalTable = Table.Sort(Table.SelectColumns(renameColumns, {"MaterialGroup", "materials"}), {{"MaterialGroup", Order.Ascending}, {"materials", Order.Ascending}})
in
    finalTable
