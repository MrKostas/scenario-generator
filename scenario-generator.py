"""Генерация файла сценария из Excel-источников в выбранной директории."""

from __future__ import annotations

import argparse
import logging
import math
import os
import re
import tempfile
from collections import defaultdict
from collections.abc import Callable
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd

DIRECTORY_PATH = Path(__file__).resolve().parent
DEFAULT_SOURCE_DIRECTORY = DIRECTORY_PATH / "data" / "raw"
OUTPUT_PATH = DIRECTORY_PATH / "Real Scenario Евдаково_новый.xlsx"
SOURCES = {
    "materials": ("DEMO-Информация о продуктах.xlsx", "materials"),
    "laboratory": ("DEMO-Лаборатория и рециркуляция.xlsx", None),
    "changeovers": ("DEMO-Переналадки.xlsx", None),
    "storage": ("DEMO-Правила хранения продуктов.xlsx", None),
    "pipes": ("DEMO-Трубы и баки.xlsx", None),
    "steps": ("DEMO-OPERATIONS", None),
    "consumable": ("DEMO-CONSUMABLES", "Consumable"),
    "consumable_record": ("DEMO-CONSUMPTION", "ConsumableRecord"),
    "stock": ("DEMO-STOCK-A", "тбх"),
    "demand": ("DEMO-DEMAND", "Итого фасовка"),
    "recipes": ("DEMO-Рецепты без Дезо и ПереЭт", None),
    "arrival": ("DEMO-ARRIVALS", "МТТ"),
    "export": ("DEMO-EXPORTS", " отгрузка с ЭПИ, Евдаково"),
    "unavailability": ("DEMO-MAINTENANCE-A", None),
    "consumable_stock": ("DEMO-PACKAGING-STOCK-A", "InitialConsumables"),
    "packaging_matrix": ("по фасовке", None),
}
EXPECTED_COLUMNS = {
    "Material": [
        "name",
        "shelfLife",
        "exportable",
        "hydrogenRate",
        "packaging",
        "SemiFinishedProduct",
        "byproduct",
        "unlimitedSupplyArea",
        "rawMaterialCategory",
        "endProductCategory",
        "technologyCategory",
        "aggregateStateCategory",
        "labStudyMode",
    ],
    "RawMaterialCategory": ["name"],
    "AggregateStateCategory": ["name"],
    "EndProductCategory": ["name"],
    "TechnologyCategory": ["name"],
    "MaterialGroup_materials": ["MaterialGroup", "materials"],
    "MaterialGroup": ["name"],
    "LabStudyMode": [
        "name",
        "testDurationType",
        "testDuration",
        "testEndTimeNextDay",
        "validityType",
        "validityPeriod",
        "triggerPolicy",
        "onExpiry",
    ],
    "WashingRecord": [
        "hourlyQuantity",
        "maxQuantity",
        "line",
        "previousMaterial",
        "nextMaterial",
        "washingMaterial",
    ],
    "ModeSwitchChangeover": ["previousMode", "nextMode", "duration", "line"],
    "MaterialTankRule": [
        "supportsDuplexMode",
        "supportsConcurrentFilling",
        "material",
        "tank",
    ],
    "Pipe": ["source", "dest"],
    "Line": ["name", "mustBeBusy", "minLoad", "lineSection"],
    "LineSection": ["name"],
    "TankGroup_tanks": ["TankGroup", "tanks"],
    "TankGroup": ["name"],
    "Tank": ["name", "capacity"],
    "PerformanceRecord": [
        "highPriorityLine",
        "mode",
        "summerPerformance",
        "winterPerformance",
        "minBatchSize",
        "batchMultiple",
        "subStep",
        "maxDailyQuantity",
        "line",
        "processingStep",
    ],
    "ComplexStep": [
        "name",
        "producedFraction1",
        "producedFraction2",
        "producedMaterial1",
        "producedMaterial2",
    ],
    "SimpleProcessingStep": ["name", "producedMaterial"],
    "Consumable": ["name", "capacityUnit"],
    "ConsumableRecord": [
        "consumedQuantity",
        "producedQuantity",
        "line",
        "step",
        "consumable",
    ],
    "InitialStock": ["mass", "productionDate", "tank", "material"],
    "LineGroup_lines": ["LineGroup", "lines"],
    "LineGroup": ["name"],
    "Scenario": [
        "name",
        "beginDate",
        "endDate",
        "firstSummerMonth",
        "firstWinterMonth",
        "pumpingRate",
    ],
    "Scenario_lineTankOutputs": ["Scenario", "lineTankOutputs"],
    "LineRestartRule": ["duration", "materialGroup", "line"],
    "LineTankOutput_lines": ["LineTankOutput", "lines"],
    "emptyChangeover": ["name", "duration", "line", "previousMaterial", "nextMaterial"],
    "packagingChangeover": [
        "name",
        "cleaningMaterialQuantity",
        "duration",
        "line",
        "previousMaterial",
        "nextMaterial",
        "cleaningMaterial",
        "changeoverMaterial",
        "changeoverMaterialTank",
        "changeoverMaterialTankGroup",
    ],
}
ACTUAL_COLUMNS = {
    "ActualArrival": ["operatingDate", "line", "material", "ID"],
    "ActualArrivalTank": [
        "mass",
        "remainingMass",
        "endDT",
        "tank",
        "operation",
        "actualArrival",
    ],
    "ActualExport": ["operatingDate", "line", "material", "ID"],
    "ActualExportTank": ["mass", "remainingMass", "endDT", "tank", "operation"],
    "ActualProduction": [
        "operatingDate",
        "endDT",
        "combinedStepName",
        "remainingQuantity",
        "ID",
    ],
    "ActualProdSourceTank": ["mass", "tank", "material", "operation"],
    "ActualProdDestTank": [
        "mass",
        "remainingLabStudyDurationHours",
        "tank",
        "material",
        "operation",
    ],
    "ActualChangeover": [
        "cleaningMaterialMass",
        "duration",
        "operatingDate",
        "line",
        "previousMaterial",
        "nextMaterial",
        "cleaningMaterial",
        "changeoverMaterial",
        "ID",
    ],
    "ActualChangeoverTank": ["mass", "endDT", "tank", "changeover"],
}
EXPECTED_COLUMNS.update(ACTUAL_COLUMNS)
OPTIONAL_COLUMNS = {
    "DemandRecord": [
        "name",
        "enabled",
        "mass",
        "beginDate",
        "endDate",
        "maxDelay",
        "priority",
        "material",
    ],
    "InputRecord": ["consumedFraction", "material", "altInput", "processingStep"],
    "AltInput": ["name", "beginDate", "endDate", "priority", "processingStep"],
    "MaterialArrivalAssignment": [
        "beginDate",
        "endDate",
        "mass",
        "productionDate",
        "material",
        "receivingArea",
    ],
    "MaterialExportAssignment": [
        "beginDate",
        "endDate",
        "mass",
        "productionDate",
        "material",
        "receivingArea",
    ],
    "UnavailabilityAssignment": ["beginDate", "endDate", "description", "equipment"],
    "MaterialSwitchChangeover": [
        "cleaningMaterialQuantity",
        "name",
        "duration",
        "line",
        "previousMaterial",
        "nextMaterial",
        "cleaningMaterial",
        "priority",
    ],
    "ForbiddenMaterialSwitch": ["line", "previousMaterial", "nextMaterial"],
    "ConsumableInitialStock": ["mass", "consumable"],
}
# Явная карта дат контракта: лист (или функция-генератор) -> поля дат.
DATE_FIELDS = {
    "Scenario": ("beginDate", "endDate"),
    "InitialStock": ("productionDate",),
    "DemandRecord": ("beginDate", "endDate"),
    "AltInput": ("beginDate", "endDate"),
    "MaterialArrivalAssignment": ("beginDate", "endDate", "productionDate"),
    "MaterialExportAssignment": ("beginDate", "endDate", "productionDate"),
    "UnavailabilityAssignment": ("beginDate", "endDate"),
}
GENERATED_DATETIME_FIELDS = {
    "InitialStock": ("productionDate",),
    "DemandRecord": ("beginDate", "endDate"),
    "MaterialArrivalAssignment": ("beginDate", "endDate", "productionDate"),
    "MaterialExportAssignment": ("beginDate", "endDate"),
    "UnavailabilityAssignment": ("beginDate", "endDate"),
}
SHEET_ORDER = [
    "Scenario",
    "Scenario_lineTankOutputs",
    "Material",
    "MaterialGroup",
    "MaterialGroup_materials",
    "SimpleProcessingStep",
    "ComplexStep",
    "PerformanceRecord",
    "AltInput",
    "InputRecord",
    "Line",
    "LineGroup",
    "LineGroup_lines",
    "Tank",
    "TankGroup",
    "TankGroup_tanks",
    "Pipe",
    "MaterialArrivalAssignment",
    "MaterialExportAssignment",
    "UnavailabilityAssignment",
    "LineRestartRule",
    "Consumable",
    "ConsumableRecord",
    "MaterialTankRule",
    "LineTankOutput_lines",
    "MaterialSwitchChangeover",
    "emptyChangeover",
    "ModeSwitchChangeover",
    "packagingChangeover",
    "ForbiddenMaterialSwitch",
    "WashingRecord",
    "InitialStock",
    "DemandRecord",
    "RawMaterialCategory",
    "EndProductCategory",
    "TechnologyCategory",
    "AggregateStateCategory",
    "LineSection",
    "ConsumableInitialStock",
    "LabStudyMode",
    *ACTUAL_COLUMNS,
]
MATERIAL_GROUPS = [
    "IBL",
    "Брендовая продукция",
    "ЗМЖ",
    "ЗМК",
    "КЖ",
    "КЖ кул",
    "КЖ простые",
    "КЖ сложные",
    "Кристаллизаторы",
    "ЛЗМК",
    "Моножиры",
    "Побочный продукт",
    "Промышленные маргарины",
    "Специализированные маргарины",
    "Фритюрные жиры",
    "ЭМК",
    "Эффективные корма",
    "Эмульгаторы",
]


@dataclass(frozen=True)
class SheetGenerationError:
    """Ошибка формирования отдельного листа сценария."""

    sheet: str
    source: str
    reason: str


class ScenarioGenerationError(Exception):
    """Накопленные ошибки, из-за которых файл сценария не был записан."""

    def __init__(self, errors: list[SheetGenerationError]) -> None:
        self.errors = errors
        super().__init__(format_generation_errors(errors))


def format_generation_errors(errors: list[SheetGenerationError]) -> str:
    """Формирует единое понятное сообщение для CLI и desktop UI."""
    return "\n\n".join(
        f"При генерации листа «{error.sheet}» произошли ошибки. "
        f"Исправьте соответствующий файл-источник «{error.source}». "
        f"Причина: {error.reason}"
        for error in errors
    )


def resolve_source_directory(source_directory: Path | None) -> Path:
    if source_directory is None:
        source_directory = (
            DEFAULT_SOURCE_DIRECTORY
            if DEFAULT_SOURCE_DIRECTORY.is_dir()
            else DIRECTORY_PATH
        )
    source_directory = source_directory.expanduser().resolve()
    if not source_directory.is_dir():
        raise NotADirectoryError(
            f"Директория источников не найдена: {source_directory}"
        )
    return source_directory


def source_path(key: str, source_directory: Path) -> Path:
    file_marker, _ = SOURCES[key]
    candidates = sorted(
        path
        for path in source_directory.rglob("*")
        if path.is_file()
        and path.suffix.lower() in {".xls", ".xlsx", ".xlsm"}
        and file_marker.casefold() in path.name.casefold()
    )

    if not candidates:
        raise FileNotFoundError(
            f"Не найден обязательный источник с кодировкой {file_marker!r}"
        )
    if len(candidates) > 1:
        names = ", ".join(path.name for path in candidates)
        raise ValueError(
            f"Для источника с кодировкой {file_marker!r} найдено несколько файлов: "
            f"{names}"
        )
    return candidates[0]


def optional_source_path(key: str, source_directory: Path) -> Path | None:
    """Возвращает единственный источник или None; неоднозначность остаётся ошибкой."""
    file_marker, _ = SOURCES[key]
    candidates = sorted(
        path
        for path in source_directory.rglob("*")
        if path.is_file()
        and path.suffix.lower() in {".xls", ".xlsx", ".xlsm"}
        and file_marker.casefold() in path.name.casefold()
    )
    if len(candidates) > 1:
        names = ", ".join(path.name for path in candidates)
        raise ValueError(
            f"Для источника с кодировкой {file_marker!r} найдено несколько файлов: {names}"
        )
    return candidates[0] if candidates else None


def read_source(
    key: str,
    source_directory: Path,
    sheet: str | None = None,
    **kwargs: object,
) -> pd.DataFrame:
    _, default_sheet = SOURCES[key]
    return pd.read_excel(
        source_path(key, source_directory), sheet_name=sheet or default_sheet, **kwargs
    )


def require_columns(data: pd.DataFrame, columns: list[str], context: str) -> None:
    missing = [column for column in columns if column not in data.columns]
    if missing:
        raise ValueError(f"{context}: отсутствуют столбцы {missing}")


def normalize_header(value: object) -> str:
    return " ".join(str(value).replace("\xa0", " ").strip().split()).casefold()


def normalize_excel_headers(
    data: pd.DataFrame, canonical_columns: list[str], context: str
) -> pd.DataFrame:
    """Сопоставляет заголовки без изменения канонического контракта результата."""
    normalized = [normalize_header(column) for column in data.columns]
    duplicates = sorted({name for name in normalized if normalized.count(name) > 1})
    if duplicates:
        raise ValueError(
            f"{context}: дублирующиеся заголовки после нормализации: {duplicates}"
        )
    canonical = {normalize_header(column): column for column in canonical_columns}
    return data.rename(
        columns={
            actual: canonical.get(name, actual)
            for actual, name in zip(data.columns, normalized, strict=True)
        }
    )


def convert_contract_dates(tables: dict[str, pd.DataFrame]) -> None:
    """Преобразует только date-поля, явно заданные в DATE_FIELDS."""
    for sheet, fields in DATE_FIELDS.items():
        if sheet not in tables:
            continue
        data = tables[sheet]
        for field in fields:
            original = data[field]
            converted = pd.to_datetime(original, dayfirst=True, errors="coerce")
            invalid = original.notna() & converted.isna()
            if invalid.any():
                raise ValueError(
                    f"{sheet}.{field}: {invalid.sum()} непустых значений не преобразованы в дату"
                )
            data[field] = converted.dt.date if sheet == "Scenario" else converted
        for field in GENERATED_DATETIME_FIELDS.get(sheet, ()):
            non_null = data[field].notna()
            if (
                non_null.any()
                and not (
                    data.loc[non_null, field].dt.time == pd.Timestamp("08:00:00").time()
                ).all()
            ):
                raise ValueError(f"{sheet}.{field}: время границы должно быть 08:00:00")


def read_contract_source(
    key: str,
    source_directory: Path,
    sheet: str | None,
    canonical_columns: list[str],
    **kwargs: object,
) -> pd.DataFrame:
    data = read_source(key, source_directory, sheet, **kwargs)
    return normalize_excel_headers(
        data, canonical_columns, f"{key}/{sheet or SOURCES[key][1]}"
    )


def category_sheet(materials: pd.DataFrame, column: str) -> pd.DataFrame:
    values = materials[column].dropna().astype("string").str.strip()
    return (
        values.loc[values.ne("")]
        .drop_duplicates()
        .rename("name")
        .to_frame()
        .reset_index(drop=True)
    )


def make_lines(pipe: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    require_columns(pipe, ["source", "dest"], "Pipe")
    names = (
        pd.concat(
            [
                pipe["source"],
                pipe["dest"],
                pd.Series(["ЦВП Электролизный", "ЦВП Метанольный"]),
            ],
            ignore_index=True,
        )
        .dropna()
        .astype("string")
        .str.strip()
    )
    names = names.loc[
        names.ne("")
        & names.str.len().gt(2)
        & ~names.str.contains(r"Бак|Вода|MAS", case=False, regex=True)
    ].drop_duplicates()
    lines = names.rename("name").to_frame().reset_index(drop=True)
    lines["mustBeBusy"] = lines["name"].str.contains(
        "Дезо", case=False, na=False
    ) & ~lines["name"].str.contains("Смеситель", case=False, na=False)
    lines["minLoad"] = None
    first_word = lines["name"].str.split().str[0]
    lines["lineSection"] = first_word.where(
        ~first_word.isin(["Зона", "Смеситель"]), "Другое"
    )
    return lines, lines["lineSection"].drop_duplicates().rename(
        "name"
    ).to_frame().reset_index(drop=True)


def make_initial_stock(
    raw_data: pd.DataFrame, planning_start: pd.Timestamp
) -> pd.DataFrame:
    marker = raw_data.iloc[:, 0].astype("string").str.strip().eq("№")
    positions = marker[marker].index
    if positions.empty:
        raise ValueError("InitialStock: не найдена строка заголовка с маркером '№'")
    stock = raw_data.iloc[positions[0] + 1 :].copy()
    stock.columns = raw_data.iloc[positions[0]].astype("string").str.strip()
    columns = ["остаток сырья, (кг)", "№ емкости", "Наименование сырья и  продукции"]
    stock = normalize_excel_headers(stock, columns, "InitialStock")
    require_columns(stock, columns, "InitialStock")
    stock["mass"] = pd.to_numeric(stock["остаток сырья, (кг)"], errors="coerce") / 1_000
    stock = stock.dropna(subset=["mass"])
    result = pd.DataFrame(
        {
            "mass": stock["mass"],
            "productionDate": planning_start,
            "tank": "Бак " + stock["№ емкости"].astype("string").str.strip(),
            "material": stock["Наименование сырья и  продукции"]
            .astype("string")
            .str.strip(),
        }
    )
    phantom = pd.DataFrame(
        {
            "mass": [1_000_000],
            "productionDate": [planning_start],
            "tank": ["Вода"],
            "material": ["Вода"],
        }
    )
    return pd.concat([result, phantom], ignore_index=True)


def make_demand_record(
    source_directory: Path,
    begin_date: pd.Timestamp,
    end_date: pd.Timestamp,
    material_groups: pd.Series | None = None,
) -> pd.DataFrame:
    """Перенос DemandRecord: итоги фасовки превращаются в заказы на месяц."""
    raw = pd.read_excel(
        source_path("demand", source_directory),
        sheet_name=SOURCES["demand"][1],
        header=None,
    )
    marker = raw.iloc[:, 0].astype("string").str.strip().eq("Итого фасовка")
    positions = marker[marker].index
    if positions.empty:
        raise ValueError("DemandRecord: не найдена строка 'Итого фасовка'")
    header = raw.iloc[positions[0] + 1]
    mass_positions = header.astype("string").str.strip().eq("Итого за месяц")
    if not mass_positions.any():
        raise ValueError("DemandRecord: не найден столбец 'Итого за месяц'")
    data = raw.iloc[positions[0] + 2 :, [0, mass_positions.idxmax()]].copy()
    data.columns = ["material", "mass"]
    material_keys = (
        data["material"]
        .astype("string")
        .str.replace(r"[\x00-\x1f]+", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .str.casefold()
    )
    data["material"] = (
        data["material"]
        .astype("string")
        .str.replace(r"[\x00-\x1f]", "", regex=True)
        .str.strip()
    )
    data["mass"] = pd.to_numeric(data["mass"], errors="coerce")

    excluded_values = ["Итого", "Исполнитель: Демо", *MATERIAL_GROUPS]
    if material_groups is not None:
        excluded_values.extend(material_groups.dropna().astype(str))
    excluded_keys = {normalize_header(value) for value in excluded_values}
    data = data.loc[
        data["material"].notna()
        & data["material"].ne("")
        & ~material_keys.isin(excluded_keys)
        & data["mass"].notna()
        & data["mass"].ne(0)
    ]
    data = data.sort_values(["mass", "material"], ascending=[False, True]).reset_index(
        drop=True
    )
    result = pd.DataFrame(
        {
            "name": "Заказ №" + (data.index + 1).astype(str),
            "enabled": True,
            "mass": data["mass"].to_numpy(),
            "beginDate": begin_date,
            "endDate": end_date,
            "maxDelay": 120,
            "priority": 1000,
            "material": data["material"].to_numpy(),
        }
    )
    return result


RECIPE_COLUMNS = ["НАИМЕНОВАНИЕ ПРОДУКЦИИ", "РЕЦЕПТ", "ПРОЦЕСС"]
REQUIRED_RECIPE_MARKERS = (
    SOURCES["recipes"][0],
    "Рецептуры Евдаково-1",
    "Рецептуры Евдаково_Эффективные корма",
)


@dataclass(frozen=True)
class RecipeRow:
    """Нормализованная рецептура с происхождением для диагностики."""

    product: str
    recipe: str
    process: str
    source: Path
    sheet: str
    row: int

    @property
    def context(self) -> str:
        return f"{self.source.name}/{self.sheet}, строка {self.row}"


def recipe_source_paths(source_directory: Path) -> list[Path]:
    """Находит все рецептуры, проверяя обязательный набор и порядок чтения."""
    paths = sorted(
        (
            path
            for path in source_directory.rglob("*")
            if path.is_file()
            and not path.name.startswith("~$")
            and path.suffix.casefold() in {".xls", ".xlsx", ".xlsm"}
            and "рецепт" in path.stem.casefold()
        ),
        key=lambda path: (
            path.relative_to(source_directory).as_posix().casefold(),
            path.relative_to(source_directory).as_posix(),
        ),
    )
    required = []
    for marker in REQUIRED_RECIPE_MARKERS:
        candidates = [
            path
            for path in paths
            if re.search(re.escape(marker) + r"(?!\d)", path.stem, re.IGNORECASE)
        ]
        if not candidates:
            raise FileNotFoundError(f"Не найден обязательный файл рецептур: {marker}")
        if len(candidates) > 1:
            raise ValueError(
                f"Для рецептур {marker!r} найдено несколько файлов: "
                + ", ".join(str(p) for p in candidates)
            )
        required.append(candidates[0])
    standard = required[0]
    return [standard, *(path for path in paths if path != standard)]


def recipe_text(value: object) -> str:
    return "" if pd.isna(value) else re.sub(r"[\x00-\x1f]", "", str(value)).strip()


def recipe_percent(text: str) -> Decimal:
    try:
        number = Decimal(text.strip().replace(",", "."))
    except InvalidOperation as error:
        raise ValueError(f"некорректный процент {text!r}") from error
    if not number.is_finite() or not 0 <= number <= 100:
        raise ValueError(f"процент должен быть конечным числом от 0 до 100: {text!r}")
    return number


def read_recipe_rows(path: Path, excluded: set[str]) -> list[RecipeRow]:
    """Читает одну книгу и приводит поддерживаемый формат к общим записям."""
    with pd.ExcelFile(path) as workbook:
        if not workbook.sheet_names:
            raise ValueError(f"{path.name}: в книге нет листов")
        is_feed = "эффективные корма" in path.stem.casefold()
        is_main = bool(
            re.search(
                re.escape(REQUIRED_RECIPE_MARKERS[1]) + r"(?!\d)",
                path.stem,
                re.IGNORECASE,
            )
        )
        is_standard = REQUIRED_RECIPE_MARKERS[0].casefold() in path.stem.casefold()
        if is_feed:
            if "Рецепт" not in workbook.sheet_names:
                raise ValueError(f"{path.name}: не найден лист Рецепт")
            sheet = "Рецепт"
        elif is_main:
            sheet = workbook.sheet_names[-1]
        elif is_standard:
            sheet = workbook.sheet_names[0]
        elif "Рецепт" in workbook.sheet_names:
            sheet = "Рецепт"
        else:
            sheet = workbook.sheet_names[0]
        raw = workbook.parse(sheet_name=sheet, header=None)
        if not (is_feed or is_main or is_standard):
            first_columns = (
                {normalize_header(value) for value in raw.iloc[0]}
                if not raw.empty
                else set()
            )
            if (
                not {normalize_header(value) for value in RECIPE_COLUMNS}.issubset(
                    first_columns
                )
                and sheet != "Рецепт"
            ):
                last_sheet = workbook.sheet_names[-1]
                if last_sheet != sheet:
                    sheet = last_sheet
                    raw = workbook.parse(sheet_name=sheet, header=None)
    context = f"{path.name}/{sheet}"
    if raw.empty:
        raise ValueError(f"{context}: пустой лист рецептур")
    markers = raw.iloc[:, 0].astype("string").str.strip().eq("N/N")
    if is_main and not markers.any():
        raise ValueError(f"{context}: не найден заголовок N/N")
    header_row = markers[markers].index[0] if markers.any() else 0
    data = raw.iloc[header_row + 1 :].copy()
    data.columns = raw.iloc[header_row]
    possible = [*RECIPE_COLUMNS, "РЕЦЕПТУРЫ ПЕР.ЖИРОВ"]
    data = normalize_excel_headers(data, possible, context)
    require_columns(data, RECIPE_COLUMNS[:2], context)
    if "ПРОЦЕСС" in data.columns:
        kind = "standard"
    elif "РЕЦЕПТУРЫ ПЕР.ЖИРОВ" in data.columns and markers.any():
        kind = "main"
    elif is_feed or "Рецепт" == sheet:
        kind = "feed"
    else:
        raise ValueError(f"{context}: неизвестный формат рецептур")
    if is_main and kind != "main":
        raise ValueError(f"{context}: отсутствует колонка РЕЦЕПТУРЫ ПЕР.ЖИРОВ")
    result = []
    section = None
    sections = {"фасовка": "Фас", "гидро": "Гидро", "смешивание": "Смеш"}
    for row_index, row in data.iterrows():
        product = recipe_text(row[RECIPE_COLUMNS[0]])
        recipe = recipe_text(row[RECIPE_COLUMNS[1]])
        pere = recipe_text(row.get("РЕЦЕПТУРЫ ПЕР.ЖИРОВ"))
        origin = f"{context}, строка {row_index + 1}"
        key = normalize_header(product)
        if kind == "feed" and key in sections and not recipe:
            section = sections[key]
            continue
        if key in excluded or not (product or recipe or pere):
            continue
        if not product:
            raise ValueError(f"{origin}: пустое наименование продукции")
        if kind == "standard":
            process = recipe_text(row["ПРОЦЕСС"])
        elif kind == "main":
            process = "Дезо"
        else:
            process = "Гран" if re.match(r"^ПФ(?:\s|$)", product) else section
        if recipe:
            if not process:
                raise ValueError(f"{origin}: не определён процесс для {product!r}")
            result.append(
                RecipeRow(product, recipe, process, path, sheet, row_index + 1)
            )
        elif not pere:
            raise ValueError(f"{origin}: отсутствует рецепт для {product!r}")
        if kind == "main" and pere:
            for variant in re.split(r"\s+/\s+", pere):
                match = re.fullmatch(r"([^%]+)%\s*([^()]+?)\s*\((.*)\)\s*", variant)
                if match is None:
                    raise ValueError(
                        f"{origin}: нераспознанная рецептура ПереЭт {variant!r}"
                    )
                try:
                    recipe_percent(match[1])
                except ValueError as error:
                    raise ValueError(f"{origin}: {error}") from error
                balance = 0
                for character in match[3]:
                    balance += (character == "(") - (character == ")")
                    if balance < 0:
                        break
                if balance:
                    raise ValueError(f"{origin}: несогласованные скобки в {variant!r}")
                if not match[2].strip():
                    raise ValueError(f"{origin}: пустой продукт в рецептуре ПереЭт")
                result.append(
                    RecipeRow(
                        match[2].strip(),
                        match[3].strip(),
                        "ПереЭт",
                        path,
                        sheet,
                        row_index + 1,
                    )
                )
    if not result:
        raise ValueError(f"{context}: не найдено ни одной рецептуры")
    return result


def recipe_reference_warnings(
    variants: dict[str, RecipeRow],
    inputs: pd.DataFrame,
    reference_tables: dict[str, pd.DataFrame],
) -> list[str]:
    """Группирует отсутствующие ссылки, не удаляя рецептуры и компоненты."""
    grouped = defaultdict(lambda: defaultdict(set))

    def missing(category: str, value: str, alternative: str):
        record = variants[alternative]
        grouped[(category, record.source.name, record.sheet)][value].add(
            f"{alternative}, строка {record.row}"
        )

    if "Material" in reference_tables:
        materials = set(reference_tables["Material"]["name"].dropna().map(recipe_text))
        for alternative, record in variants.items():
            if record.product not in materials:
                missing("продукты в Material", record.product, alternative)
        for material, alternative in (
            inputs[["material", "altInput"]]
            .drop_duplicates()
            .itertuples(index=False, name=None)
        ):
            if material not in materials:
                missing("компоненты в Material", material, alternative)
    if all(
        name in reference_tables for name in ("SimpleProcessingStep", "ComplexStep")
    ):
        steps = set(reference_tables["SimpleProcessingStep"]["name"].dropna()) | set(
            reference_tables["ComplexStep"]["name"].dropna()
        )
        for alternative, record in variants.items():
            step = f"{record.process}_{record.product}"
            if step not in steps:
                missing("операции processingStep", step, alternative)
    if "PerformanceRecord" in reference_tables:
        performance = set(
            reference_tables["PerformanceRecord"]["processingStep"].dropna()
        )
        for alternative, record in variants.items():
            step = f"{record.process}_{record.product}"
            if step not in performance:
                missing("записи производительности", step, alternative)
    return [
        f"{file}/{sheet}: отсутствуют {category}: "
        + "; ".join(
            f"{value} [{'; '.join(sorted(contexts))}]"
            for value, contexts in sorted(values.items())
        )
        for (category, file, sheet), values in sorted(grouped.items())
    ]


def make_recipe_tables(
    source_directory: Path,
    *,
    on_warning: Callable[[str], None] | None = None,
    reference_tables: dict[str, pd.DataFrame] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Объединяет все рецептуры до нумерации альтернатив и разбора компонентов."""
    excluded = {normalize_header(value) for value in ["Итого", *MATERIAL_GROUPS]}
    if reference_tables is not None and "MaterialGroup" in reference_tables:
        excluded.update(
            reference_tables["MaterialGroup"]["name"].dropna().map(normalize_header)
        )
    records = []
    for path in recipe_source_paths(source_directory):
        records.extend(read_recipe_rows(path, excluded))
    rows = []
    variants = {}
    seen = set()
    numbers = defaultdict(int)
    diagnostics = []
    for record in records:
        step = f"{record.process}_{record.product}"
        for variant in re.split(r"\s+/\s+", record.recipe):
            components = []
            try:
                for component in re.split(r"\s+\+\s+", variant):
                    percent, separator, material = component.partition("%")
                    if not separator or not material.strip():
                        raise ValueError(f"нераспознанный компонент {component!r}")
                    components.append((material.strip(), recipe_percent(percent)))
            except ValueError as error:
                raise ValueError(f"{record.context}, {step}: {error}") from error
            signature = (record.process, record.product, tuple(sorted(components)))
            if signature in seen:
                continue
            seen.add(signature)
            numbers[step] += 1
            alternative = f"{step}_Рецептура_{numbers[step]}"
            variants[alternative] = record
            total = sum((percent for _, percent in components), Decimal(0)) / 100
            if abs(total - 1) > Decimal("0.000001"):
                diagnostics.append(
                    f"{record.context}, {alternative}: сумма долей {total} вместо 1; исходные доли сохранены"
                )
            rows.extend(
                {
                    "consumedFraction": float(percent / 100),
                    "material": material,
                    "altInput": alternative,
                    "processingStep": step,
                }
                for material, percent in components
            )
    inputs = pd.DataFrame(rows, columns=OPTIONAL_COLUMNS["InputRecord"])
    alternatives = (
        inputs[["altInput", "processingStep"]]
        .drop_duplicates()
        .rename(columns={"altInput": "name"})
        .reset_index(drop=True)
    )
    alternatives["beginDate"] = pd.NaT
    alternatives["endDate"] = pd.NaT
    alternatives["priority"] = alternatives["name"].str.endswith("_Рецептура_1")
    if reference_tables is not None:
        diagnostics.extend(
            recipe_reference_warnings(variants, inputs, reference_tables)
        )
    emit_warning = (
        on_warning if on_warning is not None else logging.getLogger(__name__).warning
    )
    for message in dict.fromkeys(diagnostics):
        emit_warning(message)
    return inputs, alternatives[OPTIONAL_COLUMNS["AltInput"]]


def make_consumable_initial_stock(source_directory: Path) -> pd.DataFrame:
    data = read_source("consumable_stock", source_directory)
    data = normalize_excel_headers(
        data, ["consumables", "initialStocks"], "ConsumableInitialStock"
    )
    require_columns(data, ["consumables", "initialStocks"], "ConsumableInitialStock")
    return data.rename(columns={"consumables": "consumable", "initialStocks": "mass"})[
        OPTIONAL_COLUMNS["ConsumableInitialStock"]
    ]


def make_arrival_assignments(
    source_directory: Path, materials: pd.DataFrame, planning_start: pd.Timestamp
) -> pd.DataFrame:
    raw = read_source("arrival", source_directory, header=None)
    first = raw.iloc[:, 0].astype("string").str.casefold()
    start = first.str.contains(r"\(каменка\)", regex=True, na=False)
    end = first.str.contains(r"\(на алексеевку\)", regex=True, na=False)
    if not start.any() or not end.any():
        raise ValueError(
            "MaterialArrivalAssignment: не найдены маркеры '(Каменка)'/'(на Алексеевку)'"
        )
    section = raw.iloc[max(start.idxmax() - 1, 0) : end.idxmax()].copy()
    section.columns = section.iloc[0]
    section = section.iloc[2:].copy()
    section = section.loc[:, pd.Index(section.columns).notna()]
    section = normalize_excel_headers(
        section, ["material"], "MaterialArrivalAssignment"
    )
    material_column = section.columns[0]
    section = section.rename(columns={material_column: "material"}).drop_duplicates(
        "material"
    )
    dates = [
        column
        for column in section.columns
        if column != "material" and pd.notna(column)
    ]
    long = section.melt(
        id_vars="material",
        value_vars=dates,
        var_name="productionDate",
        value_name="mass",
    )
    long["productionDate"] = pd.to_datetime(
        long["productionDate"], dayfirst=True, errors="coerce"
    ) + pd.Timedelta(hours=8)
    long["mass"] = pd.to_numeric(long["mass"], errors="coerce")
    long = long.loc[long["mass"].notna() & long["productionDate"].notna()]
    long = long.loc[
        long["productionDate"].dt.to_period("M") == planning_start.to_period("M")
    ]
    areas = materials[["name", "unlimitedSupplyArea"]].dropna(
        subset=["unlimitedSupplyArea"]
    )
    long = long.merge(areas, left_on="material", right_on="name", how="inner").drop(
        columns="name"
    )
    long["beginDate"] = long["productionDate"] + pd.Timedelta(days=2)
    long["endDate"] = long["productionDate"] + pd.Timedelta(days=3)
    return long.rename(columns={"unlimitedSupplyArea": "receivingArea"})[
        OPTIONAL_COLUMNS["MaterialArrivalAssignment"]
    ]


def make_export_assignments(
    source_directory: Path,
    planning_start: pd.Timestamp | None = None,
    *,
    today: pd.Timestamp | None = None,
) -> pd.DataFrame:
    """Возвращает отгрузки Евдаково за месяц плана без изменения источника.

    При вызове из генератора передаётся его planning_start. Для отдельного
    расчёта используется следующий месяц относительно today (или текущей даты).
    Пустые массы пропускаются, отдельные строки баков не агрегируются.
    """
    if planning_start is not None and today is not None:
        raise ValueError("Укажите только planning_start или today")
    if planning_start is None:
        current = pd.Timestamp(today if today is not None else pd.Timestamp.today())
        planning_start = current.normalize().replace(day=1) + pd.offsets.MonthBegin(1)
    period = pd.Timestamp(planning_start).to_period("M")
    if pd.isna(period):
        raise ValueError("Некорректная дата месяца планирования")

    raw = read_source("export", source_directory, header=None)
    if raw.shape[1] < 7:
        raise ValueError("Отгрузки Евдаково: отсутствуют колонки A:G")
    labels = raw.iloc[:, 1].map(normalize_header).str.rstrip(": ")
    starts = raw.index[labels.eq("отгрузка с евдаково")].tolist()
    if len(starts) != 1:
        raise ValueError("Отгрузки Евдаково: ожидается один маркер начала секции")
    start = starts[0]
    following = labels.loc[start + 1 :]
    ends = following.index[following.str.startswith("отгрузка с ")].tolist()
    if not ends:
        raise ValueError("Отгрузки Евдаково: не найден конец секции")
    end = ends[0]
    headers = raw.index[
        labels.str.replace(" ", "", regex=False).eq("показатель/дата")
    ].tolist()
    preceding = [position for position in headers if position < start]
    if not preceding:
        raise ValueError("Отгрузки Евдаково: не найден заголовок перед секцией")
    block_headers = [preceding[-1]] + [
        position for position in headers if start < position < end
    ]
    parts = []
    calendar_found = False
    for number, header in enumerate(block_headers):
        line = raw.iat[header, 0]
        if pd.isna(line) or not str(line).strip().startswith("ПОСТ_МБХ"):
            raise ValueError(f"Отгрузки Евдаково: некорректный пост в A{header + 1}")
        dates = pd.to_datetime(raw.iloc[header, 6:], errors="coerce", dayfirst=True)
        valid_dates = dates.dropna()
        if valid_dates.empty or valid_dates.dt.normalize().duplicated().any():
            raise ValueError(
                f"Отгрузки Евдаково: некорректный календарь строки {header + 1}"
            )
        selected = dates.index[dates.dt.to_period("M").eq(period)].tolist()
        calendar_found = calendar_found or bool(selected)
        block_end = (
            block_headers[number + 1] if number + 1 < len(block_headers) else end
        )
        products = raw.iloc[max(start + 1, header + 1) : block_end]
        products = products.loc[products.iloc[:, 1].notna()]
        products = products.loc[products.iloc[:, 1].astype(str).str.strip().ne("")]
        if not selected:
            continue
        data = products.loc[:, selected].copy()
        data.columns = [dates[column].normalize() for column in selected]
        data["line"] = str(line).strip()
        data["productName"] = products.iloc[:, 1].astype(str).str.strip()
        long = data.melt(
            id_vars=["line", "productName"], var_name="date", value_name="mass"
        ).dropna(subset=["mass"])
        try:
            long["mass"] = pd.to_numeric(long["mass"], errors="raise")
        except (ValueError, TypeError) as error:
            raise ValueError(f"Отгрузки Евдаково: нечисловая масса у {line}") from error
        if not long["mass"].map(math.isfinite).all():
            raise ValueError(f"Отгрузки Евдаково: бесконечная масса у {line}")
        parts.append(long)
    if not calendar_found:
        raise ValueError(f"Отгрузки Евдаково: в источнике нет календаря за {period}")
    columns = ["line", "productName", "date", "mass"]
    result = (
        pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=columns)
    )
    result["date"] = pd.to_datetime(result["date"])
    result["mass"] = pd.to_numeric(result["mass"])
    return (
        result[columns]
        .sort_values(["date", "line"], kind="stable")
        .reset_index(drop=True)
    )


def make_export_assignment_table(
    source_directory: Path, planning_start: pd.Timestamp
) -> pd.DataFrame:
    """Приводит отгрузки к согласованному контракту файла-сценария."""
    data = make_export_assignments(source_directory, planning_start)
    data = data.rename(columns={"productName": "material", "line": "receivingArea"})
    data["beginDate"] = data["date"] + pd.Timedelta(hours=8)
    data["endDate"] = data["beginDate"] + pd.Timedelta(days=1)
    data["productionDate"] = pd.NaT
    return data[OPTIONAL_COLUMNS["MaterialExportAssignment"]]


def make_unavailability(
    source_directory: Path, planning_start: pd.Timestamp, planning_end: pd.Timestamp
) -> pd.DataFrame:
    path = source_path("unavailability", source_directory)
    rows: list[dict[str, object]] = []
    months = {
        "январь": 1,
        "февраль": 2,
        "март": 3,
        "апрель": 4,
        "май": 5,
        "июнь": 6,
        "июль": 7,
        "август": 8,
        "сентябрь": 9,
        "октябрь": 10,
        "ноябрь": 11,
        "декабрь": 12,
    }
    with pd.ExcelFile(path) as workbook:
        sheets = workbook.sheet_names
    for sheet in sheets:
        if "print_area" in sheet.casefold() or sheet.casefold() == "инструкция":
            continue
        raw = pd.read_excel(path, sheet_name=sheet, header=None)
        marker = raw.iloc[:, 0].astype("string").str.strip().str.casefold().eq("месяц")
        if not marker.any():
            continue
        index = marker.idxmax()
        equipment = str(raw.iloc[index, 1]).split("ремонта")[-1].strip() or sheet
        data = raw.iloc[index + 1 :].copy()
        data.columns = raw.iloc[index]
        month_column = data.columns[0]
        for _, record in data.iterrows():
            month = months.get(str(record[month_column]).strip().casefold())
            if not month:
                continue
            for day, description in record.iloc[1:].items():
                if pd.isna(description) or not str(description).strip():
                    continue
                day_number = pd.to_numeric(day, errors="coerce")
                if pd.isna(day_number):
                    continue
                date = pd.Timestamp(
                    year=planning_start.year, month=month, day=int(day_number), hour=8
                )
                rows.append(
                    {
                        "date": date,
                        "description": str(description).strip(),
                        "equipment": equipment,
                    }
                )
    data = pd.DataFrame(rows)
    if data.empty:
        return pd.DataFrame(columns=OPTIONAL_COLUMNS["UnavailabilityAssignment"])
    data = data.sort_values(["equipment", "description", "date"])
    data["group"] = (
        data.groupby(["equipment", "description"])["date"].diff().dt.days.ne(1)
    ).cumsum()
    result = data.groupby(["equipment", "description", "group"], as_index=False).agg(
        beginDate=("date", "min"), lastDate=("date", "max")
    )
    result["endDate"] = result["lastDate"] + pd.Timedelta(days=1)
    result = result.loc[
        (result["beginDate"] >= planning_start) & (result["beginDate"] < planning_end)
    ]
    result["equipment"] = result["equipment"].str.split(",").explode().str.strip()
    result = result.loc[~result["equipment"].isin(["ЦМБХ", "СХО ВК и СВ", "КЦ"])]
    return result[OPTIONAL_COLUMNS["UnavailabilityAssignment"]]


def make_material_switches(source_directory: Path) -> pd.DataFrame:
    """Разворачивает матрицы ЦПСЖ; фасовочные матрицы намеренно исключены."""
    agreed_lines = {
        normalize_header(
            f"Матрица переходов при смене продуктов на участке {area} ЦПСЖ"
        ): f"ЦПСЖ {area}"
        for area in ("ПереЭт и отбелка", "Дезо №2", "Дезо №3")
    }
    rows: list[dict[str, object]] = []
    for path in material_switch_matrix_paths(source_directory):
        with pd.ExcelFile(path) as workbook:
            sheets = workbook.sheet_names
        for sheet in sheets:
            if not sheet.casefold().startswith("переход_евдаково"):
                continue
            raw = pd.read_excel(path, sheet_name=sheet, header=None)
            first = raw.iloc[:, 0].astype("string")
            start = first.str.contains("с…на…", na=False)
            end = first.str.contains("промывка", case=False, na=False)
            if not start.any() or not end.any():
                continue
            title = first[
                first.map(normalize_header).str.contains(
                    "матрица переходов", regex=False
                )
            ]
            line = "ЦПСЖ " + (
                str(title.iloc[0]).split("участке")[-1].strip().split(" ")[0]
                if not title.empty
                else sheet
            )
            if not title.empty:
                line = agreed_lines.get(normalize_header(title.iloc[0]), line)
            matrix = raw.iloc[start.idxmax() : end.idxmax()].copy()
            matrix.columns = matrix.iloc[0]
            matrix = matrix.iloc[1:].ffill()
            previous_column = matrix.columns[0]
            for next_material, values in matrix.iloc[:, 1:].items():
                for previous, cleaning in zip(
                    matrix[previous_column], values, strict=True
                ):
                    if pd.isna(cleaning):
                        continue
                    for prev in str(previous).split(","):
                        for nxt in str(next_material).split(","):
                            rows.append(
                                {
                                    "duration": 1,
                                    "line": line,
                                    "previousMaterial": prev.strip(),
                                    "nextMaterial": nxt.strip(),
                                    "cleaningMaterial": str(cleaning).strip(),
                                }
                            )
    data = pd.DataFrame(rows, columns=OPTIONAL_COLUMNS["MaterialSwitchChangeover"])
    service_keys = (
        data["cleaningMaterial"].str.replace(r"\s+", " ", regex=True).str.casefold()
    )
    switches = data.loc[
        ~service_keys.isin(["без промывки", "не планировать производство"])
    ].copy()
    switches["cleaningMaterial"] = (
        switches["cleaningMaterial"]
        .str.replace(r"(?i)^промывка\s+", "", regex=True)
        .str.strip()
    )
    return switches.reset_index(drop=True)


def material_switch_matrix_paths(source_directory: Path) -> list[Path]:
    """Возвращает матрицы переходов ЦПСЖ для MaterialSwitchChangeover."""
    return [
        path
        for path in source_directory.glob("*.xls*")
        if "матрица" in path.name.casefold()
        and "цпсж" in path.name.casefold()
        and "фасовк" not in path.name.casefold()
    ]


def packaging_matrix_paths(source_directory: Path) -> list[Path]:
    """Возвращает все книги матриц фасовки в воспроизводимом порядке."""
    return sorted(
        (
            path
            for path in source_directory.rglob("*")
            if path.is_file()
            and path.suffix.casefold() in {".xls", ".xlsx", ".xlsm"}
            and not path.name.startswith("~$")
            and re.search(
                r"матриц[аы] переходов по фасовке", normalize_header(path.name)
            )
        ),
        key=lambda path: (
            path.relative_to(source_directory).as_posix().casefold(),
            path.relative_to(source_directory).as_posix(),
        ),
    )


def make_packaging_changeovers(
    source_directory: Path,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Разбирает все матрицы фасовки: время / промывочный продукт / масса."""
    paths = packaging_matrix_paths(source_directory)
    if not paths:
        raise FileNotFoundError("Не найдены Матрица/Матрицы переходов по фасовке")
    packaging_rows: list[dict[str, object]] = []
    empty_rows: list[dict[str, object]] = []

    def parse_number(value: str) -> float:
        number = float(value.strip().replace(",", "."))
        if not math.isfinite(number) or number < 0:
            raise ValueError(
                "количество и длительность должны быть конечными и неотрицательными"
            )
        return number

    for path in paths:
        try:
            workbook = pd.ExcelFile(path)
        except Exception as error:
            raise ValueError(f"{path}: не удалось открыть книгу; {error}") from error
        with workbook:
            sheets = [
                sheet
                for sheet in workbook.sheet_names
                if normalize_header(sheet).startswith("переход_евдаково")
            ]
            if not sheets:
                raise ValueError(f"{path.name}: не найдены листы Переход_Евдаково")
            for sheet in sheets:
                context = f"{path}/{sheet}"
                match = re.fullmatch(r"переход_евдаково (\d+)", normalize_header(sheet))
                if match is None or int(match.group(1)) <= 0:
                    raise ValueError(f"{context}: не определён номер линии ЦФСЖ")
                line = f"ЦФСЖ №{int(match.group(1))}"
                if workbook.engine == "openpyxl":
                    # Не разворачиваем пустую сетку по завышенным Excel dimensions.
                    worksheet = workbook.book[sheet]
                    worksheet.reset_dimensions()
                    rows = []
                    started = False
                    for row in worksheet.iter_rows(values_only=True):
                        rows.append(row)
                        first_value = normalize_header(row[0] if row else "")
                        if started and "промывка" in first_value:
                            break
                        started = started or "с…на…" in first_value
                    raw = pd.DataFrame(rows)
                else:
                    raw = workbook.parse(
                        sheet_name=sheet, header=None, keep_default_na=False
                    )
                if raw.empty:
                    raise ValueError(f"{context}: пустой лист матрицы")
                first = raw.iloc[:, 0].astype("string")
                starts = first[first.str.contains("с…на…", regex=False, na=False)].index
                if len(starts) != 1:
                    raise ValueError(f"{context}: ожидается одна строка с…на…")
                start = starts[0]
                ends = first[
                    first.str.contains("промывка", case=False, na=False)
                    & (first.index > start)
                ].index
                if ends.empty:
                    raise ValueError(f"{context}: не найдена нижняя граница матрицы")
                header = raw.iloc[start]
                columns = [
                    index
                    for index in range(1, len(header))
                    if pd.notna(header.iloc[index]) and str(header.iloc[index]).strip()
                ]
                matrix = raw.iloc[start + 1 : ends[0], [0, *columns]]
                if not columns or matrix.empty:
                    raise ValueError(f"{context}: нет продуктовых строк или столбцов")
                for row_index, previous, *values in matrix.itertuples(
                    index=True, name=None
                ):
                    for column, value in zip(columns, values, strict=True):
                        if pd.isna(value) or not str(value).strip():
                            continue
                        if normalize_header(value) in {
                            "без промывки",
                            "не планировать производство",
                        }:
                            continue
                        try:
                            if pd.isna(previous) or not str(previous).strip():
                                raise ValueError("пустое название предыдущего продукта")
                            previous_material = str(previous).strip()
                            next_material = str(header.iloc[column]).strip()
                            parts = str(value).strip().split("/")
                            if len(parts) == 1:
                                fields = {"duration": parse_number(parts[0])}
                                destination = empty_rows
                            elif len(parts) == 3 and parts[1].strip():
                                fields = {
                                    "duration": parse_number(parts[0]),
                                    "cleaningMaterial": parts[1].strip(),
                                    "cleaningMaterialQuantity": parse_number(parts[2]),
                                    "changeoverMaterial": "Смесь жиров",
                                    "changeoverMaterialTank": None,
                                    "changeoverMaterialTankGroup": "Баки Смеси жиров",
                                }
                                destination = packaging_rows
                            else:
                                raise ValueError(
                                    "ожидается число или время/материал/масса"
                                )
                        except ValueError as error:
                            raise ValueError(
                                f"{context}, ячейка R{row_index + 1}C{column + 1}: {value!r}; {error}"
                            ) from error
                        name = f"Переход_{previous_material}_{next_material}"
                        if "cleaningMaterial" in fields:
                            name += f"_{fields['cleaningMaterial']}"
                        destination.append(
                            {
                                "name": name,
                                "line": line,
                                "previousMaterial": previous_material,
                                "nextMaterial": next_material,
                                **fields,
                            }
                        )
    return (
        pd.DataFrame(packaging_rows, columns=EXPECTED_COLUMNS["packagingChangeover"]),
        pd.DataFrame(empty_rows, columns=EXPECTED_COLUMNS["emptyChangeover"]),
    )


def validate_tables(tables: dict[str, pd.DataFrame]) -> None:
    for sheet, columns in EXPECTED_COLUMNS.items():
        if sheet not in tables:
            raise ValueError(f"Не сформирован обязательный лист {sheet}")
        if list(tables[sheet].columns) != columns:
            raise ValueError(f"{sheet}: структура не соответствует эталону")
    for sheet, columns in OPTIONAL_COLUMNS.items():
        if sheet in tables and list(tables[sheet].columns) != columns:
            raise ValueError(f"{sheet}: структура не соответствует Power Query")


def add_sheet(
    tables: dict[str, pd.DataFrame],
    errors: list[SheetGenerationError],
    sheet: str,
    source: str,
    producer: Callable[[], pd.DataFrame],
) -> pd.DataFrame | None:
    """Собирает один лист и сохраняет ошибку, не прерывая остальные шаги."""
    try:
        data = producer()
        columns = EXPECTED_COLUMNS.get(sheet)
        if columns is None:
            columns = OPTIONAL_COLUMNS[sheet]
        require_columns(data, columns, sheet)
        data = data[columns]
        convert_contract_dates({sheet: data})
    except Exception as error:  # noqa: BLE001
        errors.append(SheetGenerationError(sheet, source, str(error)))
        return None
    tables[sheet] = data
    return data


def build_tables(
    today: pd.Timestamp | None = None,
    source_directory: Path | None = None,
    *,
    on_warning: Callable[[str], None] | None = None,
) -> dict[str, pd.DataFrame]:
    source_directory = resolve_source_directory(source_directory)
    current = (today or pd.Timestamp.today()).normalize()
    next_month = current.replace(day=1) + pd.offsets.MonthBegin(1)
    planning_start = next_month + pd.Timedelta(hours=8)
    planning_end = next_month + pd.offsets.MonthBegin(1) + pd.Timedelta(hours=8)
    tables: dict[str, pd.DataFrame] = {}
    errors: list[SheetGenerationError] = []

    def source_name(key: str) -> str:
        return SOURCES[key][0]

    def optional_available(sheets: tuple[str, ...], key: str) -> bool:
        try:
            available = optional_source_path(key, source_directory) is not None
        except Exception as error:  # noqa: BLE001
            for sheet in sheets:
                errors.append(SheetGenerationError(sheet, source_name(key), str(error)))
            return False
        if not available:
            for sheet in sheets:
                errors.append(
                    SheetGenerationError(
                        sheet,
                        source_name(key),
                        "не найден обязательный файл-источник",
                    )
                )
        return available

    def dependency_error(sheet: str, source: str, prerequisite: str) -> None:
        errors.append(
            SheetGenerationError(
                sheet,
                source,
                f"не сформирован, так как не сформирован зависимый лист «{prerequisite}»",
            )
        )

    material_columns = [
        column
        for column in EXPECTED_COLUMNS["Material"]
        if column not in ("packaging", "SemiFinishedProduct", "byproduct")
    ]

    def make_materials() -> pd.DataFrame:
        materials = read_contract_source(
            "materials", source_directory, None, material_columns
        )
        require_columns(materials, material_columns, "Material")
        materials["packaging"] = materials["technologyCategory"].isin(
            ["Отгруженные наливом", "Фасованные жиры"]
        )
        materials["SemiFinishedProduct"] = None
        materials["byproduct"] = None
        return materials

    materials = add_sheet(
        tables, errors, "Material", source_name("materials"), make_materials
    )
    if materials is not None:
        for sheet, column in (
            ("RawMaterialCategory", "rawMaterialCategory"),
            ("AggregateStateCategory", "aggregateStateCategory"),
            ("EndProductCategory", "endProductCategory"),
            ("TechnologyCategory", "technologyCategory"),
        ):
            add_sheet(
                tables,
                errors,
                sheet,
                source_name("materials"),
                lambda column=column: category_sheet(materials, column),
            )

        group_materials = add_sheet(
            tables,
            errors,
            "MaterialGroup_materials",
            source_name("materials"),
            lambda: materials.loc[
                materials["technologyCategory"].isin(
                    ["Дезодорированные жиры", "Переэтерифицированные жиры"]
                ),
                ["endProductCategory", "name"],
            ]
            .rename(
                columns={"endProductCategory": "MaterialGroup", "name": "materials"}
            )
            .sort_values(["MaterialGroup", "materials"], na_position="last")
            .reset_index(drop=True),
        )
        if group_materials is not None:
            add_sheet(
                tables,
                errors,
                "MaterialGroup",
                source_name("materials"),
                lambda: pd.concat(
                    [
                        group_materials["MaterialGroup"],
                        pd.Series(["Эффективные корма", "Эмульгаторы"]),
                    ]
                )
                .drop_duplicates()
                .rename("name")
                .to_frame()
                .reset_index(drop=True),
            )
        else:
            dependency_error(
                "MaterialGroup", source_name("materials"), "MaterialGroup_materials"
            )
    else:
        for sheet in (
            "RawMaterialCategory",
            "AggregateStateCategory",
            "EndProductCategory",
            "TechnologyCategory",
            "MaterialGroup_materials",
            "MaterialGroup",
        ):
            dependency_error(sheet, source_name("materials"), "Material")

    for sheet in ("LabStudyMode", "WashingRecord"):
        add_sheet(
            tables,
            errors,
            sheet,
            source_name("laboratory"),
            lambda sheet=sheet: read_contract_source(
                "laboratory", source_directory, sheet, EXPECTED_COLUMNS[sheet]
            ),
        )
    add_sheet(
        tables,
        errors,
        "ModeSwitchChangeover",
        source_name("changeovers"),
        lambda: read_contract_source(
            "changeovers",
            source_directory,
            "modeChange",
            EXPECTED_COLUMNS["ModeSwitchChangeover"],
        ),
    )
    add_sheet(
        tables,
        errors,
        "ForbiddenMaterialSwitch",
        source_name("changeovers"),
        lambda: read_contract_source(
            "changeovers",
            source_directory,
            "ForbiddenMaterialSwitch",
            OPTIONAL_COLUMNS["ForbiddenMaterialSwitch"],
        ),
    )
    add_sheet(
        tables,
        errors,
        "MaterialTankRule",
        source_name("storage"),
        lambda: read_contract_source(
            "storage",
            source_directory,
            "хранениеПродуктов",
            EXPECTED_COLUMNS["MaterialTankRule"],
        ),
    )

    pipe = add_sheet(
        tables,
        errors,
        "Pipe",
        source_name("pipes"),
        lambda: read_contract_source(
            "pipes", source_directory, "трубы", EXPECTED_COLUMNS["Pipe"]
        ),
    )
    tank_groups = add_sheet(
        tables,
        errors,
        "TankGroup_tanks",
        source_name("pipes"),
        lambda: read_contract_source(
            "pipes",
            source_directory,
            "группыБаков",
            EXPECTED_COLUMNS["TankGroup_tanks"],
        ),
    )
    add_sheet(
        tables,
        errors,
        "Tank",
        source_name("pipes"),
        lambda: read_contract_source(
            "pipes", source_directory, "баки", EXPECTED_COLUMNS["Tank"]
        ),
    )
    if pipe is not None:
        line_data = add_sheet(
            tables, errors, "Line", source_name("pipes"), lambda: make_lines(pipe)[0]
        )
        if line_data is not None:
            add_sheet(
                tables,
                errors,
                "LineSection",
                source_name("pipes"),
                lambda: line_data["lineSection"]
                .drop_duplicates()
                .rename("name")
                .to_frame()
                .reset_index(drop=True),
            )
        else:
            dependency_error("LineSection", source_name("pipes"), "Line")
    else:
        dependency_error("Line", source_name("pipes"), "Pipe")
        dependency_error("LineSection", source_name("pipes"), "Pipe")
    if tank_groups is not None:
        add_sheet(
            tables,
            errors,
            "TankGroup",
            source_name("pipes"),
            lambda: tank_groups["TankGroup"]
            .drop_duplicates()
            .rename("name")
            .to_frame()
            .reset_index(drop=True),
        )
    else:
        dependency_error("TankGroup", source_name("pipes"), "TankGroup_tanks")

    for sheet in ("PerformanceRecord", "ComplexStep", "SimpleProcessingStep"):
        add_sheet(
            tables,
            errors,
            sheet,
            source_name("steps"),
            lambda sheet=sheet: read_contract_source(
                "steps",
                source_directory,
                sheet,
                [
                    column
                    for column in EXPECTED_COLUMNS[sheet]
                    if column != "maxDailyQuantity"
                ],
            ).assign(
                **({"maxDailyQuantity": None} if sheet == "PerformanceRecord" else {})
            ),
        )
    add_sheet(
        tables,
        errors,
        "Consumable",
        source_name("consumable"),
        lambda: read_contract_source(
            "consumable", source_directory, None, EXPECTED_COLUMNS["Consumable"]
        ),
    )
    add_sheet(
        tables,
        errors,
        "ConsumableRecord",
        source_name("consumable_record"),
        lambda: read_contract_source(
            "consumable_record",
            source_directory,
            None,
            EXPECTED_COLUMNS["ConsumableRecord"],
        ),
    )
    add_sheet(
        tables,
        errors,
        "InitialStock",
        source_name("stock"),
        lambda: make_initial_stock(
            read_source("stock", source_directory, header=None), planning_start
        ),
    )

    for sheet, data in {
        "LineGroup_lines": pd.DataFrame(columns=EXPECTED_COLUMNS["LineGroup_lines"]),
        "LineGroup": pd.DataFrame(columns=EXPECTED_COLUMNS["LineGroup"]),
        "Scenario": pd.DataFrame(
            {
                "name": ["scenario name"],
                "beginDate": [next_month.date()],
                "endDate": [(next_month + pd.offsets.MonthEnd(0)).date()],
                "firstSummerMonth": [5],
                "firstWinterMonth": [10],
                "pumpingRate": [25],
            }
        ),
        "Scenario_lineTankOutputs": pd.DataFrame(
            columns=EXPECTED_COLUMNS["Scenario_lineTankOutputs"]
        ),
        "LineRestartRule": pd.DataFrame(columns=EXPECTED_COLUMNS["LineRestartRule"]),
        "LineTankOutput_lines": pd.DataFrame(
            columns=EXPECTED_COLUMNS["LineTankOutput_lines"]
        ),
    }.items():
        add_sheet(tables, errors, sheet, "Генератор сценариев", lambda data=data: data)

    if optional_available(("DemandRecord",), "demand"):
        add_sheet(
            tables,
            errors,
            "DemandRecord",
            source_name("demand"),
            lambda: make_demand_record(
                source_directory,
                planning_start,
                planning_end,
                material_groups=(
                    materials["endProductCategory"] if materials is not None else None
                ),
            ),
        )
    try:
        input_record, alt_input = make_recipe_tables(
            source_directory,
            on_warning=on_warning,
            reference_tables=tables,
        )
    except Exception as error:  # noqa: BLE001
        for sheet in ("InputRecord", "AltInput"):
            errors.append(
                SheetGenerationError(sheet, "книги рецептур Евдаково", str(error))
            )
    else:
        add_sheet(
            tables,
            errors,
            "InputRecord",
            "книги рецептур Евдаково",
            lambda: input_record,
        )
        add_sheet(
            tables, errors, "AltInput", "книги рецептур Евдаково", lambda: alt_input
        )
    if optional_available(("MaterialArrivalAssignment",), "arrival"):
        if materials is None:
            dependency_error(
                "MaterialArrivalAssignment", source_name("arrival"), "Material"
            )
        else:
            add_sheet(
                tables,
                errors,
                "MaterialArrivalAssignment",
                source_name("arrival"),
                lambda: make_arrival_assignments(
                    source_directory, materials, planning_start
                ),
            )
    if optional_available(("MaterialExportAssignment",), "export"):
        add_sheet(
            tables,
            errors,
            "MaterialExportAssignment",
            source_name("export"),
            lambda: make_export_assignment_table(source_directory, planning_start),
        )
    if optional_available(("UnavailabilityAssignment",), "unavailability"):
        add_sheet(
            tables,
            errors,
            "UnavailabilityAssignment",
            source_name("unavailability"),
            lambda: make_unavailability(source_directory, planning_start, planning_end),
        )
    if optional_available(("ConsumableInitialStock",), "consumable_stock"):
        add_sheet(
            tables,
            errors,
            "ConsumableInitialStock",
            source_name("consumable_stock"),
            lambda: make_consumable_initial_stock(source_directory),
        )
    if not material_switch_matrix_paths(source_directory):
        errors.append(
            SheetGenerationError(
                "MaterialSwitchChangeover",
                "матрица переходов ЦПСЖ",
                "не найден обязательный файл-источник",
            )
        )
    else:
        try:
            switches = make_material_switches(source_directory)
        except Exception as error:  # noqa: BLE001
            errors.append(
                SheetGenerationError(
                    "MaterialSwitchChangeover", "матрица переходов ЦПСЖ", str(error)
                )
            )
        else:
            add_sheet(
                tables,
                errors,
                "MaterialSwitchChangeover",
                "матрица переходов ЦПСЖ",
                lambda: switches,
            )

    try:
        packaging, empty = make_packaging_changeovers(source_directory)
    except Exception as error:  # noqa: BLE001
        for sheet in ("packagingChangeover", "emptyChangeover"):
            errors.append(
                SheetGenerationError(sheet, source_name("packaging_matrix"), str(error))
            )
    else:
        for sheet, data in (
            ("packagingChangeover", packaging),
            ("emptyChangeover", empty),
        ):
            add_sheet(
                tables,
                errors,
                sheet,
                source_name("packaging_matrix"),
                lambda data=data: data,
            )

    for sheet, columns in ACTUAL_COLUMNS.items():
        add_sheet(
            tables,
            errors,
            sheet,
            "Генератор сценариев",
            lambda columns=columns: pd.DataFrame(columns=columns),
        )

    if errors:
        raise ScenarioGenerationError(errors)
    tables = {sheet: tables[sheet] for sheet in SHEET_ORDER if sheet in tables}
    validate_tables(tables)
    return tables


def write_scenario(
    tables: dict[str, pd.DataFrame], output_path: Path = OUTPUT_PATH
) -> None:
    with tempfile.NamedTemporaryFile(
        prefix="scenario_", suffix=".xlsx", dir=output_path.parent, delete=False
    ) as temp_file:
        temp_path = Path(temp_file.name)
    try:
        with pd.ExcelWriter(
            temp_path,
            engine="openpyxl",
            date_format="DD.MM.YYYY",
            datetime_format="DD.MM.YYYY HH:MM:SS",
        ) as writer:
            for sheet, data in tables.items():
                data.to_excel(writer, sheet_name=sheet, index=False)
                if sheet == "Scenario":
                    for field in DATE_FIELDS["Scenario"]:
                        column = data.columns.get_loc(field) + 1
                        for row in writer.sheets[sheet].iter_rows(
                            min_row=2,
                            max_row=len(data) + 1,
                            min_col=column,
                            max_col=column,
                        ):
                            row[0].number_format = "DD.MM.YYYY"
        os.replace(temp_path, output_path)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def create_scenario(
    source_directory: Path | None = None,
    output_path: Path | None = None,
    *,
    on_warning: Callable[[str], None] | None = None,
) -> Path:
    """Создаёт книгу сценария из источников и возвращает путь к результату."""
    output_path = Path(output_path) if output_path is not None else OUTPUT_PATH
    if on_warning is None:
        tables = build_tables(source_directory=source_directory)
    else:
        tables = build_tables(source_directory=source_directory, on_warning=on_warning)
    write_scenario(tables, output_path)
    return output_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Генератор файла сценария")
    parser.add_argument(
        "--source-dir",
        type=Path,
        help=(
            "Путь к директории с Excel-источниками (абсолютный или относительный). "
            "Без параметра используется data/raw."
        ),
    )
    parser.add_argument(
        "--output-path",
        type=Path,
        help="Путь к создаваемому файлу сценария.",
    )
    arguments = parser.parse_args()
    source_directory = resolve_source_directory(arguments.source_dir)
    print("Создан файл: " f"{create_scenario(source_directory, arguments.output_path)}")
