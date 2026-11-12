"""Create synthetic Excel inputs from the public ETL contracts."""

from __future__ import annotations

import argparse
import sys
from datetime import date, datetime, timezone
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from desktop_app import SITE_CONFIGURATIONS, load_generator_module

MONTHS = (
    "январь",
    "февраль",
    "март",
    "апрель",
    "май",
    "июнь",
    "июль",
    "август",
    "сентябрь",
    "октябрь",
    "ноябрь",
    "декабрь",
)
OIL = "Масло Демо"
PRODUCT = "Продукт Демо"
PACK = "Упаковка Демо"
AREA = "Приёмка Демо"
LINE = "Линия Демо"
TANK = "Бак D01"


def write_book(directory, filename, sheets):
    path = directory / filename
    if path.exists():
        raise FileExistsError(path)
    book = Workbook()
    book.remove(book.active)
    book.properties.creator = "Scenario Generator demo"
    for name, rows in sheets.items():
        sheet = book.create_sheet(name)
        for row in rows:
            sheet.append(row)
    book.save(path)


def table(headers, **values):
    return [headers, [values.get(column) for column in headers]]


def create_inputs(directory: Path, site: str, as_of: date, months: int = 1) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    configuration = SITE_CONFIGURATIONS[site]
    pipeline = load_generator_module(configuration.generator_path)
    epi = site == "ЭПИ"
    schema = pipeline.SCHEMAS if epi else pipeline.EXPECTED_COLUMNS
    planning_month = (
        pd.Timestamp(as_of).replace(day=1) + pd.offsets.MonthBegin()
    ).date()
    movement_days = [
        (pd.Timestamp(planning_month) + pd.offsets.MonthBegin(index))
        .date()
        .replace(day=3)
        for index in range(months)
    ]

    def form(key, sheets):
        marker, _ = pipeline.SOURCES[key]
        filename = marker if marker.endswith(".xlsx") else f"{marker}.xlsx"
        write_book(directory, filename, sheets)

    def direct(sheet, **values):
        return table(schema[sheet], **values)

    material_columns = [
        column
        for column in schema["Material"]
        if column not in {"packaging", "SemiFinishedProduct", "byproduct"}
    ]
    oil = {
        "name": OIL,
        "shelfLife": 30,
        "exportable": False,
        "rawMaterialCategory": "Демо-сырьё",
        "endProductCategory": "Группа Демо",
        "technologyCategory": "Дезодорированные жиры",
        "aggregateStateCategory": "Жидкость",
        "unlimitedSupplyArea": AREA,
    }
    product = {
        "name": PRODUCT,
        "shelfLife": 60,
        "exportable": True,
        "endProductCategory": "Группа Демо",
        "technologyCategory": "Фасованные жиры",
        "aggregateStateCategory": "Твёрдое",
        "unlimitedSupplyArea": AREA,
    }
    materials = table(material_columns, **oil)
    materials.append(table(material_columns, **product)[1])
    steps = [(f"Фас_{PRODUCT}", PRODUCT)]
    if not epi:
        for process, name in (("Дезо", "Демо Дезо"), ("Фас", "Демо Фас")):
            materials.append(table(material_columns, **{**product, "name": name})[1])
            steps.append((f"{process}_{name}", name))
    form("materials", {"materials": materials})
    form(
        "laboratory",
        {
            "LabStudyMode": direct("LabStudyMode", name="Лаборатория Демо"),
            "WashingRecord": direct("WashingRecord", line=LINE, washingMaterial=OIL),
        },
    )
    changeovers = {
        "modeChange": direct(
            "ModeSwitchChangeover",
            previousMode="A",
            nextMode="B",
            duration=2,
            line=LINE,
        )
    }
    if not epi:
        changeovers["ForbiddenMaterialSwitch"] = [
            pipeline.OPTIONAL_COLUMNS["ForbiddenMaterialSwitch"]
        ]
    form("changeovers", changeovers)
    form(
        "storage",
        {"хранениеПродуктов": direct("MaterialTankRule", material=OIL, tank=TANK)},
    )
    pipes = {
        "трубы": direct("Pipe", source=TANK, dest=LINE),
        "группыБаков": direct("TankGroup_tanks", TankGroup="Группа Демо", tanks=TANK),
    }
    if not epi:
        pipes["баки"] = direct("Tank", name=TANK, capacity=15)
    form("pipes", pipes)
    form(
        "performance" if epi else "steps",
        {
            "PerformanceRecord": [
                schema["PerformanceRecord"],
                *(
                    direct("PerformanceRecord", line=LINE, processingStep=step)[1]
                    for step, _ in steps
                ),
            ],
            "ComplexStep": [schema["ComplexStep"]],
            "SimpleProcessingStep": [
                schema["SimpleProcessingStep"],
                *(
                    direct("SimpleProcessingStep", name=step, producedMaterial=name)[1]
                    for step, name in steps
                ),
            ],
        },
    )

    form(
        "demand",
        {
            "Итого фасовка": [
                ["Итого фасовка"],
                ["Продукт", "Итого за месяц"],
                [PRODUCT, 42],
                ["Группа Демо", 9],
                ["Итого", 51],
            ]
        },
    )
    section = ("(Алексеевка)", "(Каменка)") if epi else ("(Каменка)", "(на Алексеевку)")
    arrival = [["material", *movement_days], [section[0]]]
    if epi:
        arrival.append(["Заголовок секции"])
    arrival.extend([[OIL, *([12.5] * months)], [section[1]]])
    form("arrival", {"МТТ": arrival})

    recipe_headers = ["НАИМЕНОВАНИЕ ПРОДУКЦИИ", "РЕЦЕПТ"]
    form(
        "recipes",
        {"Рецепт": [[*recipe_headers, "ПРОЦЕСС"], [PRODUCT, f"100% {OIL}", "Фас"]]},
    )
    matrix = [["с…на…", PRODUCT], [OIL, f"Промывка {OIL}"]]
    repairs = [["Месяц", f"График ремонта {LINE}", 4, 5]]
    if epi:
        repairs.append(["Календарь", None, 4, 5])
    repairs.extend([month, None, "Демо", "Демо"] for month in MONTHS)
    form("unavailability", {"Ремонты": repairs})

    if epi:
        form(
            "production_recipes",
            {"Рецепты": [["N/N", *recipe_headers, "РЕЦЕПТУРЫ ПЕР.ЖИРОВ"]]},
        )
        form(
            "initial_stock",
            {
                "Остатки": [
                    [
                        None,
                        "№ п/п",
                        "Группа емкостей",
                        "Наименование сырья/продукции",
                        "№ емкости",
                        "Масса, кг",
                    ],
                    [None, 1, "Демо", OIL, "D01", 2500],
                ]
            },
        )
        form(
            "consumable_movement",
            {
                "DEMO-CONSUMABLE-MOVEMENT": [
                    [None, "Наименование", "Ед. измерения"],
                    [1, PACK, "шт"],
                    [None, "Диспетчер"],
                ]
            },
        )
        form(
            "packaging_stock",
            {
                "Остатки ": [
                    ["№ п/п"],
                    ["Номер", "Наименование", "Единица", "Остаток", PRODUCT],
                    [None, None, None, None, "норма"],
                    [1, PACK, "шт", 18, 3],
                ]
            },
        )
        for key in ("initial_consumables_1", "initial_consumables_2"):
            form(
                key,
                {
                    pipeline.SOURCES[key][1]: [
                        [None, "№ п/п", "Материал", "Показатель", "Итого"],
                        [None, 1, PACK, "Ост.факт.по справке", None, 18],
                        [None, None, "Исполнитель: Демо"],
                    ]
                },
            )
        write_book(
            directory,
            "Параметры МБХ Демо.xlsx",
            {
                "Баки": [
                    [
                        "№ ёмкости",
                        "Кол-во заполнения при плотности 0,898 г/смᶟ (NAOH плотности 1,45 г/смᶟ)",
                    ],
                    ["D01", 15000],
                ]
            },
        )
        write_book(
            directory,
            "Матрица Демо.xlsx",
            {
                "Дезо_Алекс": [
                    ["Матрица переходов на участке 1 Демо"],
                    *matrix,
                    ["При недействительной записи"],
                ]
            },
        )
    else:
        form(
            "consumable",
            {"Consumable": direct("Consumable", name=PACK, capacityUnit="шт")},
        )
        form(
            "consumable_record",
            {
                "ConsumableRecord": direct(
                    "ConsumableRecord",
                    consumable=PACK,
                    consumedQuantity=3,
                    producedQuantity=1,
                    line=LINE,
                )
            },
        )
        form(
            "stock",
            {
                "тбх": [
                    [
                        "№",
                        "остаток сырья, (кг)",
                        "№ емкости",
                        "Наименование сырья и  продукции",
                    ],
                    [1, 2500, "D01", OIL],
                ]
            },
        )
        form(
            "consumable_stock",
            {"InitialConsumables": [["consumables", "initialStocks"], [PACK, 18]]},
        )
        form(
            "export",
            {
                " отгрузка с ЭПИ, Евдаково": [
                    [
                        "ПОСТ_МБХ Демо",
                        "Показатель/дата",
                        None,
                        None,
                        None,
                        None,
                        *movement_days,
                    ],
                    [None, "Отгрузка с Евдаково"],
                    [None, PRODUCT, None, None, None, None, *([7.5] * months)],
                    [None, "Отгрузка с другой площадки"],
                ]
            },
        )
        write_book(
            directory,
            "Рецептуры Евдаково-1.xlsx",
            {
                "Рецептуры": [
                    ["N/N", *recipe_headers, "РЕЦЕПТУРЫ ПЕР.ЖИРОВ"],
                    [1, "Демо Дезо", f"100% {OIL}"],
                ]
            },
        )
        write_book(
            directory,
            "Рецептуры Евдаково_Эффективные корма.xlsx",
            {"Рецепт": [recipe_headers, ["Фасовка"], ["Демо Фас", f"100% {OIL}"]]},
        )
        write_book(
            directory,
            "Матрица переходов ЦПСЖ Демо.xlsx",
            {
                "Переход_Евдаково 1": [
                    ["Матрица переходов на участке Демо ЦПСЖ"],
                    *matrix,
                    ["Промывка"],
                ]
            },
        )
        write_book(
            directory,
            "Матрица переходов по фасовке Демо.xlsx",
            {
                "Переход_Евдаково 1": [
                    ["с…на…", PRODUCT, OIL],
                    [OIL, f"2 / {OIL} / 0,5", "1"],
                    ["Промывка"],
                ]
            },
        )


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    local_date = datetime.now(timezone.utc).astimezone().date()
    parser.add_argument("--as-of", type=date.fromisoformat, default=local_date)
    parser.add_argument("--months", type=int, default=1)
    args = parser.parse_args()
    if not 1 <= args.months <= 120:
        parser.error("--months must be between 1 and 120")
    for site, folder in (("Евдаково", "site-a"), ("ЭПИ", "site-b")):
        create_inputs(args.output / folder, site, args.as_of, args.months)


if __name__ == "__main__":
    main()
