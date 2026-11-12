"""Генератор сценария для площадки ЭПИ.

Модуль намеренно отделён от ``scenario-generator.py``: у ЭПИ другие
источники, листы и правила разбора. Правила ниже перенесены из
``power-query/epi``; при schema drift книга не записывается.
"""

from __future__ import annotations

import argparse
import os
import tempfile
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

DIRECTORY_PATH = Path(__file__).resolve().parent
DEFAULT_SOURCE_DIRECTORY = DIRECTORY_PATH / "data" / "raw" / "ЭПИ"
OUTPUT_PATH = DIRECTORY_PATH / "Real Scenario ЭПИ.xlsx"

# Маркеры и имена листов взяты из power-query/epi и M-запросов.
SOURCES = {
    "materials": ("DEMO-Информация о продуктах", "materials"),
    "laboratory": ("DEMO-Лаборатория и рециркуляция", None),
    "changeovers": ("DEMO-Переналадки", None),
    "storage": ("DEMO-Правила хранения продуктов", None),
    "pipes": ("DEMO-Трубы и баки", None),
    "performance": ("DEMO-PERFORMANCE", None),
    "consumable_movement": ("DEMO-CONSUMABLE-MOVEMENT", None),
    "packaging_stock": ("DEMO-PACKAGING-STOCK-B", None),
    "initial_stock": ("DEMO-STOCK-B", "Остатки"),
    "demand": ("DEMO-DEMAND", "Итого фасовка"),
    "recipes": ("DEMO-Рецепты без Дезо и ПереЭт", None),
    "production_recipes": ("Рецептуры ЭПИ", None),
    "arrival": ("DEMO-ARRIVALS", "МТТ"),
    "unavailability": ("DEMO-MAINTENANCE-B", None),
    "initial_consumables_1": ("DEMO-CONSUMABLE-BALANCE", "DEMO-CONSUMABLE-BALANCE"),
    "initial_consumables_2": ("DEMO-CONSUMABLE-BALANCE(1)", "DEMO-PERFORMANCE"),
}

SCHEMAS = {
    "Scenario": ["name", "beginDate", "endDate", "firstSummerMonth", "firstWinterMonth", "pumpingRate"],
    "Material": ["name", "shelfLife", "exportable", "hydrogenRate", "packaging", "unlimitedSupplyArea", "rawMaterialCategory", "endProductCategory", "technologyCategory", "aggregateStateCategory", "labStudyMode"],
    "RawMaterialCategory": ["name"],
    "AggregateStateCategory": ["name"],
    "EndProductCategory": ["name"],
    "TechnologyCategory": ["name"],
    "MaterialGroup_materials": ["MaterialGroup", "materials"],
    "MaterialGroup": ["name"],
    "LabStudyMode": ["name", "testDurationType", "testDuration", "testEndTimeNextDay", "validityType", "validityPeriod", "triggerPolicy", "onExpiry"],
    "WashingRecord": ["line", "previousMaterial", "nextMaterial", "washingMaterial", "hourlyQuantity", "maxQuantity"],
    "ModeSwitchChangeover": ["previousMode", "nextMode", "duration", "line"],
    "MaterialTankRule": ["supportsDuplexMode", "supportsConcurrentFilling", "material", "tank"],
    "Pipe": ["source", "dest"],
    "Line": ["name", "mustBeBusy", "lineSection"],
    "LineSection": ["name"],
    "TankGroup_tanks": ["TankGroup", "tanks"],
    "TankGroup": ["name"],
    "Tank": ["name", "capacity"],
    "PerformanceRecord": ["highPriorityLine", "mode", "summerPerformance", "winterPerformance", "minBatchSize", "batchMultiple", "subStep", "line", "processingStep"],
    "ComplexStep": ["name", "producedFraction1", "producedFraction2", "producedMaterial1", "producedMaterial2"],
    "SimpleProcessingStep": ["name", "producedMaterial"],
    "Consumable": ["name", "capacityUnit"],
    "ConsumableRecord": ["consumedQuantity", "producedQuantity", "line", "step", "consumable"],
    "InitialStock": ["mass", "productionDate", "tank", "material"],
    "ConsumableInitialStock": ["consumables", "initialStocks", "beginDate"],
    "DemandRecord": ["name", "enabled", "mass", "beginDate", "endDate", "maxDelay", "priority", "material"],
    "InputRecord": ["consumedFraction", "material", "altInput", "processingStep"],
    "AltInput": ["name", "beginDate", "endDate", "priority", "processingStep"],
    "MaterialArrivalAssignment": ["beginDate", "endDate", "mass", "productionDate", "material", "receivingArea"],
    "UnavailabilityAssignment": ["beginDate", "endDate", "description", "equipment"],
    "LineGroup_lines": ["LineGroup", "lines"],
    "LineGroup": ["name"],
    "MaterialSwitchChangeover": ["duration", "line", "previousMaterial", "nextMaterial", "cleaningMaterial"],
    "ForbiddenMaterialSwitch": ["line", "previousMaterial", "nextMaterial"],
    "Scenario_lineTankOutputs": ["Scenario", "lineTankOutputs"],
    "LineRestartRule": ["duration", "materialGroup", "line"],
    "LineTankOutput_lines": ["LineTankOutput", "lines"],
    "emptyChangeover": ["name", "duration", "line", "previousMaterial", "nextMaterial"],
    "packagingChangeover": ["name", "cleaningMaterialQuantity", "duration", "line", "previousMaterial", "nextMaterial", "cleaningMaterial", "changeoverMaterial", "changeoverMaterialTank", "changeoverMaterialTankGroup"],
}
SHEET_ORDER = [
    "Scenario", "Scenario_lineTankOutputs", "Material", "MaterialGroup", "MaterialGroup_materials",
    "SimpleProcessingStep", "ComplexStep", "PerformanceRecord", "AltInput",
    "InputRecord", "Line", "LineGroup", "LineGroup_lines", "Tank",
    "TankGroup", "TankGroup_tanks", "Pipe", "MaterialArrivalAssignment",
    "UnavailabilityAssignment", "LineRestartRule", "Consumable", "ConsumableRecord",
    "MaterialTankRule", "LineTankOutput_lines", "MaterialSwitchChangeover",
    "emptyChangeover", "ModeSwitchChangeover", "packagingChangeover",
    "ForbiddenMaterialSwitch", "WashingRecord", "InitialStock", "DemandRecord",
    "RawMaterialCategory", "EndProductCategory", "TechnologyCategory",
    "AggregateStateCategory", "LineSection", "ConsumableInitialStock", "LabStudyMode",
]
# Нерешённые контракты: пять листов пока имеют только пустую структуру,
# MaterialExportAssignment отсутствует и в книге Евдаково, а схема остатков ВСМ
# отличается от Евдаково при одинаковом названии листа.
KNOWN_BLOCKERS = {
    "MaterialExportAssignment": "в EPI нет M-спецификации и схемы листа",
    "Scenario_lineTankOutputs": "пустая схема; в EPI нет источника и правила формирования",
    "LineRestartRule": "пустая схема; в EPI нет источника и правила формирования",
    "LineTankOutput_lines": "пустая схема; в EPI нет источника и правила формирования",
    "emptyChangeover": "согласованная пустая таблица: матрица фасовки ЭПИ отсутствует",
    "packagingChangeover": "согласованная пустая таблица: матрица фасовки ЭПИ отсутствует",
    "ConsumableInitialStock": "имя листа совпадает с Евдаково, но EPI-спецификация формирует поля consumables, initialStocks, beginDate",
}
DATE_FIELDS = {
    "Scenario": ("beginDate", "endDate"),
    "InitialStock": ("productionDate",),
    "ConsumableInitialStock": ("beginDate",),
    "DemandRecord": ("beginDate", "endDate"),
    "AltInput": ("beginDate", "endDate"),
    "MaterialArrivalAssignment": ("beginDate", "endDate", "productionDate"),
    "UnavailabilityAssignment": ("beginDate", "endDate"),
}
GENERATED_DATETIME_FIELDS = {
    sheet: fields for sheet, fields in DATE_FIELDS.items() if sheet != "AltInput"
}


@dataclass(frozen=True)
class SheetGenerationError:
    sheet: str
    source: str
    reason: str


class ScenarioGenerationError(Exception):
    def __init__(self, errors: list[SheetGenerationError]) -> None:
        self.errors = errors
        super().__init__(format_generation_errors(errors))


def format_generation_errors(errors: list[SheetGenerationError]) -> str:
    return "\n\n".join(
        f"При генерации листа «{error.sheet}» произошли ошибки. "
        f"Исправьте соответствующий файл-источник «{error.source}». "
        f"Причина: {error.reason}"
        for error in errors
    )


def normalize_header(value: object) -> str:
    return " ".join(str(value).replace("\xa0", " ").strip().split()).casefold()


def require_columns(data: pd.DataFrame, columns: list[str], context: str) -> None:
    missing = [column for column in columns if column not in data.columns]
    if missing:
        raise ValueError(f"{context}: отсутствуют столбцы {missing}")


def normalize_excel_headers(data: pd.DataFrame, canonical: list[str], context: str) -> pd.DataFrame:
    normalized = [normalize_header(column) for column in data.columns]
    expected = {normalize_header(column) for column in canonical}
    duplicates = sorted({name for name in normalized if name in expected and normalized.count(name) > 1})
    if duplicates:
        raise ValueError(f"{context}: дублирующиеся заголовки после нормализации: {duplicates}")
    names = {normalize_header(column): column for column in canonical}
    return data.rename(columns={actual: names.get(name, actual) for actual, name in zip(data.columns, normalized, strict=True)})


def resolve_source_directory(source_directory: Path | None) -> Path:
    resolved = (source_directory or DEFAULT_SOURCE_DIRECTORY).expanduser().resolve()
    if not resolved.is_dir():
        raise NotADirectoryError(f"Директория источников не найдена: {resolved}")
    return resolved


def source_path(key: str, source_directory: Path) -> Path:
    marker, _ = SOURCES[key]
    candidates = sorted(
        path for path in source_directory.rglob("*")
        if path.is_file() and path.suffix.lower() in {".xls", ".xlsx", ".xlsm"}
        and marker.casefold() in path.name.casefold()
    )
    if key == "initial_consumables_1":
        candidates = [path for path in candidates if "(1)" not in path.name]
    if not candidates:
        raise FileNotFoundError(f"Не найден обязательный источник с кодировкой {marker!r}")
    if len(candidates) != 1:
        raise ValueError(f"Для источника с кодировкой {marker!r} найдено несколько файлов: {', '.join(path.name for path in candidates)}")
    return candidates[0]


def read_source(key: str, source_directory: Path, sheet: str | int | None = None, **kwargs: object) -> pd.DataFrame:
    _, default_sheet = SOURCES[key]
    return pd.read_excel(source_path(key, source_directory), sheet_name=default_sheet if sheet is None else sheet, **kwargs)


def read_contract_source(key: str, source_directory: Path, sheet: str | int | None, columns: list[str]) -> pd.DataFrame:
    data = read_source(key, source_directory, sheet)
    data = normalize_excel_headers(data, columns, f"{key}/{sheet}")
    require_columns(data, columns, f"{key}/{sheet}")
    return data[columns]


def category_sheet(materials: pd.DataFrame, column: str) -> pd.DataFrame:
    values = materials[column].dropna().astype("string").str.strip()
    return values.loc[values.ne("")].drop_duplicates().rename("name").to_frame().reset_index(drop=True)


def make_lines(pipe: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    names = pd.concat([pipe["source"], pipe["dest"], pd.Series(["ЦВП Электролизный", "ЦВП Метанольный"])], ignore_index=True)
    names = names.dropna().astype("string").str.strip()
    names = names.loc[names.ne("") & names.str.len().gt(2) & ~names.str.contains(r"Бак|MAS", case=False, regex=True) & names.ne("Вода")].drop_duplicates()
    lines = names.rename("name").to_frame().reset_index(drop=True)
    lines["mustBeBusy"] = lines["name"].str.contains("Дезо", case=False, na=False) & ~lines["name"].str.contains("Смеситель", case=False, na=False)
    first_word = lines["name"].str.split().str[0]
    lines["lineSection"] = first_word.where(~first_word.isin(["Зона", "Смеситель"]), "Другое")
    return lines, lines["lineSection"].drop_duplicates().rename("name").to_frame().reset_index(drop=True)


def first_marker_row(raw: pd.DataFrame, column: int, marker: str, context: str) -> int:
    values = raw.iloc[:, column].astype("string").str.strip()
    positions = values[values.str.casefold().eq(marker.casefold())].index
    if positions.empty:
        raise ValueError(f"{context}: не найден маркер {marker!r}")
    return int(positions[0])


def stock_table_from_raw(raw: pd.DataFrame, planning_start: pd.Timestamp) -> pd.DataFrame:
    header_row = first_marker_row(raw, 1, "№ п/п", "InitialStock")
    data = raw.iloc[header_row + 1 :].copy()
    data.columns = raw.iloc[header_row]
    columns = ["Группа емкостей", "Наименование сырья/продукции", "№ емкости", "Масса, кг"]
    data = normalize_excel_headers(data, columns, "InitialStock")
    require_columns(data, columns, "InitialStock")
    data["Масса, кг"] = pd.to_numeric(data["Масса, кг"], errors="coerce")
    data = data.loc[data["Масса, кг"].notna() & data["Масса, кг"].ne(0)].copy()
    data[["Группа емкостей"]] = data[["Группа емкостей"]].ffill()
    tank = data["№ емкости"].astype("string").str.replace("DEMO/", " ", regex=False).str.strip()
    result = pd.DataFrame({"mass": (data["Масса, кг"] / 1_000).round(5), "productionDate": planning_start, "tank": "Бак " + tank, "material": data["Наименование сырья/продукции"].astype("string").str.strip()})
    phantom = pd.DataFrame({"mass": [1_000_000, 1_000_000, 10_000_000], "productionDate": [planning_start] * 3, "tank": ["MAS E6", "MAS E7", "Вода"], "material": ["MASESTER E6000", "MASESTER E7000", "Вода"]})
    return pd.concat([result, phantom], ignore_index=True)


def packaging_stock_table(raw: pd.DataFrame, context: str) -> pd.DataFrame:
    header_row = first_marker_row(raw, 0, "№ п/п", context)
    data = raw.iloc[header_row + 1 :].copy()
    data.columns = raw.iloc[header_row + 1]
    data = data.iloc[1:].copy()
    data = data.loc[data.iloc[:, 0].notna()].copy()
    return data


def make_consumable(source_directory: Path) -> pd.DataFrame:
    raw_1 = read_source("consumable_movement", source_directory, "DEMO-CONSUMABLE-MOVEMENT", header=None)
    start = first_marker_row(raw_1, 1, "Наименование", "Consumable/DEMO-CONSUMABLE-MOVEMENT")
    end_mask = raw_1.iloc[:, 1].astype("string").str.casefold().str.contains("диспетчер", na=False)
    ends = end_mask[end_mask].index
    if ends.empty or ends[0] <= start:
        raise ValueError("Consumable/DEMO-CONSUMABLE-MOVEMENT: не найдена граница «диспетчер»")
    first = raw_1.iloc[start + 1 : ends[0]].copy(); first.columns = raw_1.iloc[start]
    first = normalize_excel_headers(first, ["Наименование", "Ед. измерения"], "Consumable/DEMO-CONSUMABLE-MOVEMENT")
    require_columns(first, ["Наименование", "Ед. измерения"], "Consumable/DEMO-CONSUMABLE-MOVEMENT")
    first = first[["Наименование", "Ед. измерения"]].rename(columns={"Наименование": "name", "Ед. измерения": "capacityUnit"})
    second = packaging_stock_table(read_source("packaging_stock", source_directory, "Остатки ", header=None), "Consumable/DEMO-PACKAGING-STOCK-B")
    if second.shape[1] < 3:
        raise ValueError("Consumable/DEMO-PACKAGING-STOCK-B: требуется не менее трёх столбцов")
    second = second.iloc[:, [1, 2]].copy()
    second.columns = ["name", "capacityUnit"]
    result = pd.concat([first, second], ignore_index=True)
    result = result.dropna(subset=["name"]).astype({"name": "string", "capacityUnit": "string"})
    result["name"] = result["name"].str.strip(); result["capacityUnit"] = result["capacityUnit"].str.strip()
    return result.loc[result["name"].ne("")].drop_duplicates("name").reset_index(drop=True)


def make_consumable_records(source_directory: Path) -> pd.DataFrame:
    raw = read_source("packaging_stock", source_directory, "Остатки ", header=None)
    marker = first_marker_row(raw, 0, "№ п/п", "ConsumableRecord")
    products = raw.iloc[marker + 1]
    units = raw.iloc[marker + 2].astype("string").str.strip().str.casefold()
    positions = [index for index in range(4, raw.shape[1]) if pd.notna(products.iloc[index]) and units.iloc[index] == "норма"]
    if not positions:
        raise ValueError("ConsumableRecord: не найдены столбцы нормы расхода")
    data = raw.iloc[marker + 3 :].loc[lambda frame: frame.iloc[:, 0].notna()]
    consumable = data.iloc[:, 1].ffill().astype("string").str.strip()
    long = pd.concat(
        [pd.DataFrame({"consumable": consumable, "attribute": str(products.iloc[index]).strip(), "consumedQuantity": pd.to_numeric(data.iloc[:, index], errors="coerce")}) for index in positions],
        ignore_index=True,
    )
    long = long.loc[long["consumedQuantity"].notna()].copy()
    long["step"] = long["attribute"].where(long["attribute"].str.contains("налив", case=False, na=False), "Фас_" + long["attribute"])
    long.loc[long["attribute"].str.contains("налив", case=False, na=False), "step"] = "Слив_" + long.loc[long["attribute"].str.contains("налив", case=False, na=False), "attribute"]
    long["producedQuantity"] = 1
    long["line"] = (["ЦФСЖ №1", "ЦФСЖ №2", "ЦФСЖ №3"] * ((len(long) + 2) // 3))[: len(long)]
    long["consumedQuantity"] = long["consumedQuantity"].round(5)
    return long[SCHEMAS["ConsumableRecord"]].reset_index(drop=True)


def make_initial_consumables(source_directory: Path, planning_start: pd.Timestamp) -> pd.DataFrame:
    parts: list[pd.DataFrame] = []
    for key in ("initial_consumables_1", "initial_consumables_2"):
        raw = read_source(key, source_directory, header=None)
        start = first_marker_row(raw, 1, "№ п/п", f"InitialConsumables/{key}")
        end_mask = raw.iloc[:, 2].astype("string").str.contains("Исполнитель:", na=False)
        ends = end_mask[end_mask].index
        if ends.empty or ends[0] <= start:
            raise ValueError(f"InitialConsumables/{key}: не найдена граница «Исполнитель:»")
        data = raw.iloc[start + 1 : ends[0]].copy(); data.columns = raw.iloc[start]
        total_positions = [index for index, column in enumerate(data.columns) if normalize_header(column) == "итого"]
        consumable_position = 2
        if not total_positions or consumable_position + 1 >= data.shape[1]:
            raise ValueError(f"InitialConsumables/{key}: не найдены обязательные столбцы")
        result = data.iloc[:, [consumable_position, consumable_position + 1, total_positions[0] + 1]].copy()
        result.columns = ["consumables", "attribute", "initialStocks"]
        result["consumables"] = result["consumables"].ffill().astype("string").str.strip()
        result = result.loc[result["attribute"].astype("string").str.strip().eq("Ост.факт.по справке")].copy()
        result["initialStocks"] = pd.to_numeric(result["initialStocks"], errors="coerce").fillna(0).round(5)
        result["beginDate"] = planning_start
        parts.append(result[SCHEMAS["ConsumableInitialStock"]])
    return pd.concat(parts, ignore_index=True)


def make_demand(source_directory: Path, begin_date: pd.Timestamp, end_date: pd.Timestamp, material_groups: pd.Series) -> pd.DataFrame:
    raw = read_source("demand", source_directory, header=None)
    header_row = first_marker_row(raw, 0, "Итого фасовка", "DemandRecord")
    header = raw.iloc[header_row + 1]
    mass_positions = header.astype("string").str.strip().eq("Итого за месяц")
    if not mass_positions.any():
        raise ValueError("DemandRecord: не найден столбец «Итого за месяц»")
    data = raw.iloc[header_row + 2 :, [0, mass_positions.idxmax()]].copy(); data.columns = ["material", "mass"]
    material_keys = (
        data["material"]
        .astype("string")
        .str.replace(r"[\x00-\x1f]+", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .str.casefold()
    )
    data["material"] = data["material"].astype("string").str.replace(r"[\x00-\x1f]", "", regex=True).str.strip()
    data["mass"] = pd.to_numeric(data["mass"], errors="coerce")
    data = data.loc[data["material"].notna() & data["material"].ne("") & data["mass"].notna() & data["mass"].ne(0)]
    excluded = set(material_groups.dropna().astype(str)) | {
        "Итого",
        "Исполнитель: Демо",
        "КЖ",
        "ЗМЖ",
        "ЗМК",
    }
    excluded_keys = {normalize_header(value) for value in excluded}
    data = data.loc[~material_keys.loc[data.index].isin(excluded_keys)].reset_index(
        drop=True
    )
    return pd.DataFrame({"name": "Заказ №" + (data.index + 1).astype(str), "enabled": True, "mass": data["mass"], "beginDate": begin_date, "endDate": end_date, "maxDelay": 120, "priority": 1000, "material": data["material"]})


def make_recipe_tables(source_directory: Path, material_groups: pd.Series) -> tuple[pd.DataFrame, pd.DataFrame]:
    path = source_path("production_recipes", source_directory)
    with pd.ExcelFile(path) as workbook:
        if not workbook.sheet_names:
            raise ValueError("Рецептуры: в книге нет листов")
        raw = workbook.parse(sheet_name=workbook.sheet_names[-1], header=None)
    marker = first_marker_row(raw, 0, "N/N", "Рецептуры")
    data = raw.iloc[marker + 1 :].copy(); data.columns = raw.iloc[marker]
    columns = ["НАИМЕНОВАНИЕ ПРОДУКЦИИ", "РЕЦЕПТ", "РЕЦЕПТУРЫ ПЕР.ЖИРОВ"]
    data = normalize_excel_headers(data, columns, "Рецептуры")
    require_columns(data, columns, "Рецептуры")
    excluded = set(material_groups.dropna().astype(str)) | {"Итого"}
    base = data.loc[~data["НАИМЕНОВАНИЕ ПРОДУКЦИИ"].isin(excluded), columns].copy()
    dezo = base[["НАИМЕНОВАНИЕ ПРОДУКЦИИ", "РЕЦЕПТ"]].dropna(); dezo["ПРОЦЕСС"] = "Дезо"
    pere_rows: list[dict[str, str]] = []
    for text in base["РЕЦЕПТУРЫ ПЕР.ЖИРОВ"].dropna().astype(str):
        for variant in text.split(" / "):
            if "%" not in variant or "(" not in variant or ")" not in variant:
                raise ValueError(f"Рецептуры ПЕР.ЖИРОВ: нераспознанная запись {variant!r}")
            product = variant.split("%", 1)[1].strip().split(" ", 1)[0]
            pere_rows.append({"НАИМЕНОВАНИЕ ПРОДУКЦИИ": product, "РЕЦЕПТ": variant.split("(", 1)[1].rsplit(")", 1)[0], "ПРОЦЕСС": "ПереЭт"})
    standard = read_contract_source("recipes", source_directory, 0, ["НАИМЕНОВАНИЕ ПРОДУКЦИИ", "РЕЦЕПТ", "ПРОЦЕСС"])
    recipes = pd.concat([dezo, pd.DataFrame(pere_rows), standard], ignore_index=True).dropna(subset=["НАИМЕНОВАНИЕ ПРОДУКЦИИ", "РЕЦЕПТ", "ПРОЦЕСС"])
    rows: list[dict[str, object]] = []
    for product, recipe, process in recipes.itertuples(index=False, name=None):
        step = f"{str(process).strip()}_{str(product).strip()}"
        for number, variant in enumerate(str(recipe).split(" / "), start=1):
            alternative = f"{step}_Рецептура_{number}"
            for component in variant.split(" + "):
                percent, separator, material = component.partition("%")
                fraction = pd.to_numeric(percent.replace(",", "."), errors="coerce")
                if not separator or pd.isna(fraction):
                    raise ValueError(f"Рецептуры: нераспознанный компонент {component!r}")
                rows.append({"consumedFraction": round(float(fraction) / 100, 3), "material": material.strip(), "altInput": alternative, "processingStep": step})
    inputs = pd.DataFrame(rows, columns=SCHEMAS["InputRecord"])
    alternatives = inputs[["altInput", "processingStep"]].drop_duplicates().rename(columns={"altInput": "name"})
    alternatives["beginDate"] = pd.NaT; alternatives["endDate"] = pd.NaT
    alternatives["priority"] = alternatives["name"].str.endswith("_Рецептура_1")
    return inputs, alternatives[SCHEMAS["AltInput"]]


def make_arrivals(source_directory: Path, materials: pd.DataFrame, month_start: pd.Timestamp) -> pd.DataFrame:
    raw = read_source("arrival", source_directory, header=None)
    first = raw.iloc[:, 0].astype("string").str.casefold()
    start = first.str.contains(r"\(алексеевка\)", regex=True, na=False); end = first.str.contains(r"\(каменка\)", regex=True, na=False)
    if not start.any() or not end.any() or end.idxmax() <= start.idxmax():
        raise ValueError("MaterialArrivalAssignment: не найдены границы «(Алексеевка)»/«(Каменка)»")
    section = raw.iloc[max(start.idxmax() - 1, 0) : end.idxmax()].copy(); section.columns = section.iloc[0]; section = section.iloc[3:].copy()
    section = section.loc[:, pd.Index(section.columns).notna()]
    material_column = section.columns[0]; section = section.rename(columns={material_column: "material"}).drop_duplicates("material")
    dates = [column for column in section.columns if column != "material" and pd.notna(column)]
    long = section.melt(id_vars="material", value_vars=dates, var_name="productionDate", value_name="mass")
    long["productionDate"] = pd.to_datetime(long["productionDate"], dayfirst=True, errors="coerce") + pd.Timedelta(hours=8)
    long["mass"] = pd.to_numeric(long["mass"], errors="coerce")
    long = long.loc[long["productionDate"].notna() & long["mass"].notna() & long["productionDate"].dt.to_period("M").eq(month_start.to_period("M"))]
    long = long.merge(materials[["name", "unlimitedSupplyArea"]].dropna(), left_on="material", right_on="name", how="inner").drop(columns="name")
    long["beginDate"] = long["productionDate"] + pd.Timedelta(days=2); long["endDate"] = long["productionDate"] + pd.Timedelta(days=3)
    return long.rename(columns={"unlimitedSupplyArea": "receivingArea"}).assign(mass=lambda frame: frame["mass"].round(5))[SCHEMAS["MaterialArrivalAssignment"]]


def make_unavailability(source_directory: Path, month_start: pd.Timestamp, month_end: pd.Timestamp) -> pd.DataFrame:
    months = {"январь": 1, "февраль": 2, "март": 3, "апрель": 4, "май": 5, "июнь": 6, "июль": 7, "август": 8, "сентябрь": 9, "октябрь": 10, "ноябрь": 11, "декабрь": 12}
    excluded = {"МБХ", "ЦГРМ", "ЦПВ №1", "ЦПВ №2", "ЦФСЖ участок №2 на 2026 год"}
    rows: list[dict[str, object]] = []
    path = source_path("unavailability", source_directory)
    with pd.ExcelFile(path) as workbook:
        sheets = workbook.sheet_names
    for sheet in sheets:
        raw = pd.read_excel(path, sheet_name=sheet, header=None)
        marker = raw.iloc[:, 0].astype("string").str.strip().str.casefold().eq("месяц")
        if not marker.any():
            continue
        index = marker.idxmax(); equipment = str(raw.iloc[index, 1]).split("ремонта")[-1].strip() or sheet
        data = raw.iloc[index + 2:].copy(); data.columns = raw.iloc[index + 1]
        for _, record in data.iterrows():
            month = months.get(str(record.iloc[0]).strip().casefold())
            if not month:
                continue
            for day, description in record.iloc[1:].items():
                day_number = pd.to_numeric(day, errors="coerce")
                if pd.isna(description) or not str(description).strip() or pd.isna(day_number):
                    continue
                rows.append({"date": pd.Timestamp(month_start.year, month, int(day_number), 8), "description": str(description).strip(), "equipment": equipment})
    data = pd.DataFrame(rows)
    if data.empty:
        return pd.DataFrame(columns=SCHEMAS["UnavailabilityAssignment"])
    data = data.sort_values(["equipment", "description", "date"]); data["group"] = (data.groupby(["equipment", "description"])["date"].diff().dt.days.ne(1)).cumsum()
    result = data.groupby(["equipment", "description", "group"], as_index=False).agg(beginDate=("date", "min"), lastDate=("date", "max"))
    result["endDate"] = result["lastDate"] + pd.Timedelta(days=1); result = result.loc[(result["beginDate"] >= month_start) & (result["beginDate"] < month_end)]
    result = result.assign(equipment=result["equipment"].str.split(",")).explode("equipment")
    result["equipment"] = result["equipment"].str.strip()
    result = result.loc[~result["equipment"].isin(excluded)]
    return result[SCHEMAS["UnavailabilityAssignment"]]


def make_line_groups(lines: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    selected = lines.loc[lines["name"].str.contains("МСлив", case=False, na=False), ["name"]].copy()
    if selected.empty:
        return pd.DataFrame(columns=SCHEMAS["LineGroup_lines"]), pd.DataFrame(columns=SCHEMAS["LineGroup"])
    selected["base"] = selected["name"].str.replace(r"\s+№[^ ]+$", "", regex=True)
    special = selected["name"].str.contains(r"МСлив 2_Авто №[34]", case=False, regex=True, na=False)
    normal_bases = selected.loc[~special, "base"].drop_duplicates().reset_index(drop=True)
    mapping = {base: f"Зона приемки_{index + 1}" for index, base in normal_bases.items()}
    normal = selected.loc[~special, ["name", "base"]].assign(LineGroup=lambda frame: frame["base"].map(mapping))
    special_rows = selected.loc[special, ["name"]].loc[lambda frame: frame.index.repeat(2)].copy(); special_rows["LineGroup"] = ["ЦФСЖ Авто №6", "ЦФСЖ Авто №7"] * (len(special_rows) // 2)
    result = pd.concat([normal[["LineGroup", "name"]], special_rows[["LineGroup", "name"]]], ignore_index=True)
    groups = result["LineGroup"].drop_duplicates().rename("name").to_frame().reset_index(drop=True)
    result = result.rename(columns={"name": "lines"})
    return result, groups


def matrix_paths(source_directory: Path) -> list[Path]:
    return sorted(path for path in source_directory.rglob("*.xls*") if "матрица" in path.name.casefold())


def make_material_switches(source_directory: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    rows: list[dict[str, object]] = []
    for path in matrix_paths(source_directory):
        raw = pd.read_excel(path, sheet_name="Дезо_Алекс", header=None)
        first = raw.iloc[:, 0].astype("string")
        starts = first.str.contains("с…на…", na=False); ends = first.str.contains("недействительной", case=False, na=False)
        titles = first.str.contains("Матрица переходов", case=False, na=False)
        if not starts.any() or not ends.any() or not titles.any():
            raise ValueError(f"{path.name}: не найдены маркеры матрицы переходов")
        title = str(first[titles].iloc[0]); tail = title.split("участке", 1)[-1].strip(); line = f"{tail.split(' ', 1)[-1]} Дезо"
        matrix = raw.iloc[starts.idxmax() + 1 : ends.idxmax()].copy(); matrix.columns = raw.iloc[starts.idxmax()]; matrix = matrix.loc[:, ~pd.Index(matrix.columns).astype("string").str.contains("Column", case=False)].ffill()
        previous_column = matrix.columns[0]
        for next_material, values in matrix.iloc[:, 1:].items():
            for previous, cleaning in zip(matrix[previous_column], values, strict=True):
                if pd.isna(cleaning):
                    continue
                for previous_part in str(previous).split(","):
                    for next_part in str(next_material).split(","):
                        rows.append(
                            {
                                "duration": 1,
                                "line": line,
                                "previousMaterial": previous_part.strip(),
                                "nextMaterial": next_part.strip(),
                                "cleaningMaterial": str(cleaning).strip(),
                            }
                        )
    data = pd.DataFrame(rows, columns=SCHEMAS["MaterialSwitchChangeover"])
    service_keys = (
        data["cleaningMaterial"].str.replace(r"\s+", " ", regex=True).str.casefold()
    )
    forbidden = data.loc[
        service_keys.eq("не планировать производство"),
        SCHEMAS["ForbiddenMaterialSwitch"],
    ]
    switches = data.loc[
        ~service_keys.isin(["без промывки", "не планировать производство"])
    ].copy()
    switches["cleaningMaterial"] = (
        switches["cleaningMaterial"]
        .str.replace(r"(?i)^промывка\s+", "", regex=True)
        .str.strip()
    )
    return switches.reset_index(drop=True), forbidden.reset_index(drop=True)


def convert_dates(sheet: str, data: pd.DataFrame) -> None:
    for field in DATE_FIELDS.get(sheet, ()):
        original = data[field]; converted = pd.to_datetime(original, dayfirst=True, errors="coerce")
        if (original.notna() & converted.isna()).any():
            raise ValueError(f"{sheet}.{field}: непустые значения не преобразованы в дату")
        data[field] = converted
        non_null = converted.notna()
        if (
            field in GENERATED_DATETIME_FIELDS.get(sheet, ())
            and non_null.any()
            and not (converted.loc[non_null].dt.time == pd.Timestamp("08:00:00").time()).all()
        ):
            raise ValueError(f"{sheet}.{field}: время границы должно быть 08:00:00")


def add_sheet(tables: dict[str, pd.DataFrame], errors: list[SheetGenerationError], sheet: str, source: str, producer: Callable[[], pd.DataFrame]) -> pd.DataFrame | None:
    try:
        data = producer(); require_columns(data, SCHEMAS[sheet], sheet); data = data[SCHEMAS[sheet]].copy(); convert_dates(sheet, data)
    except Exception as error:  # noqa: BLE001
        errors.append(SheetGenerationError(sheet, source, str(error))); return None
    tables[sheet] = data; return data


def build_tables(today: pd.Timestamp | None = None, source_directory: Path | None = None) -> dict[str, pd.DataFrame]:
    source_directory = resolve_source_directory(source_directory)
    current = (today or pd.Timestamp.today()).normalize(); month_start = current.replace(day=1) + pd.offsets.MonthBegin(1); month_end = month_start + pd.offsets.MonthBegin(1)
    boundary = month_start + pd.Timedelta(hours=8); end_boundary = month_end + pd.Timedelta(hours=8)
    tables: dict[str, pd.DataFrame] = {}; errors: list[SheetGenerationError] = []
    name = lambda key: SOURCES[key][0]
    materials = add_sheet(tables, errors, "Material", name("materials"), lambda: read_contract_source("materials", source_directory, "materials", [column for column in SCHEMAS["Material"] if column != "packaging"]).assign(packaging=lambda data: data["technologyCategory"].isin(["Отгруженные наливом", "Фасованные жиры"])))
    if materials is not None:
        for sheet, column in (("RawMaterialCategory", "rawMaterialCategory"), ("AggregateStateCategory", "aggregateStateCategory"), ("EndProductCategory", "endProductCategory"), ("TechnologyCategory", "technologyCategory")):
            add_sheet(tables, errors, sheet, name("materials"), lambda column=column: category_sheet(materials, column))
        groups = add_sheet(tables, errors, "MaterialGroup_materials", name("materials"), lambda: materials.loc[materials["technologyCategory"].isin(["Дезодорированные жиры", "Переэтерифицированные жиры"]), ["endProductCategory", "name"]].rename(columns={"endProductCategory": "MaterialGroup", "name": "materials"}).sort_values(["MaterialGroup", "materials"]).reset_index(drop=True))
        if groups is not None:
            add_sheet(
                tables,
                errors,
                "MaterialGroup",
                name("materials"),
                lambda: pd.concat(
                    [
                        groups["MaterialGroup"],
                        pd.Series(
                            ["Кристаллизаторы", "Эффективные корма", "Эмульгаторы"]
                        ),
                    ]
                )
                .drop_duplicates()
                .rename("name")
                .to_frame()
                .reset_index(drop=True),
            )
    for sheet in ("LabStudyMode", "WashingRecord"):
        add_sheet(tables, errors, sheet, name("laboratory"), lambda sheet=sheet: read_contract_source("laboratory", source_directory, sheet, SCHEMAS[sheet]))
    add_sheet(tables, errors, "ModeSwitchChangeover", name("changeovers"), lambda: read_contract_source("changeovers", source_directory, "modeChange", SCHEMAS["ModeSwitchChangeover"]))
    add_sheet(tables, errors, "MaterialTankRule", name("storage"), lambda: read_contract_source("storage", source_directory, "хранениеПродуктов", SCHEMAS["MaterialTankRule"]))
    pipe = add_sheet(tables, errors, "Pipe", name("pipes"), lambda: read_contract_source("pipes", source_directory, "трубы", SCHEMAS["Pipe"]))
    tank_groups = add_sheet(tables, errors, "TankGroup_tanks", name("pipes"), lambda: read_contract_source("pipes", source_directory, "группыБаков", SCHEMAS["TankGroup_tanks"]))
    add_sheet(tables, errors, "Tank", "параметры мбх", lambda: make_tanks(source_directory))
    if pipe is not None:
        lines = add_sheet(tables, errors, "Line", name("pipes"), lambda: make_lines(pipe)[0])
        if lines is not None:
            add_sheet(tables, errors, "LineSection", name("pipes"), lambda: make_lines(pipe)[1])
            line_groups, groups = make_line_groups(lines); add_sheet(tables, errors, "LineGroup_lines", "Line", lambda: line_groups); add_sheet(tables, errors, "LineGroup", "Line", lambda: groups)
    if tank_groups is not None:
        add_sheet(
            tables,
            errors,
            "TankGroup",
            name("pipes"),
            lambda: tank_groups["TankGroup"]
            .drop_duplicates()
            .rename("name")
            .to_frame()
            .reset_index(drop=True),
        )
    for sheet in ("PerformanceRecord", "ComplexStep", "SimpleProcessingStep"):
        add_sheet(tables, errors, sheet, name("performance"), lambda sheet=sheet: read_contract_source("performance", source_directory, sheet, SCHEMAS[sheet]))
    add_sheet(tables, errors, "Consumable", "DEMO-CONSUMABLE-MOVEMENT и DEMO-PACKAGING-STOCK-B", lambda: make_consumable(source_directory))
    add_sheet(tables, errors, "ConsumableRecord", name("packaging_stock"), lambda: make_consumable_records(source_directory))
    add_sheet(tables, errors, "InitialStock", name("initial_stock"), lambda: stock_table_from_raw(read_source("initial_stock", source_directory, header=None), boundary))
    add_sheet(tables, errors, "ConsumableInitialStock", "DEMO-CONSUMABLE-BALANCE", lambda: make_initial_consumables(source_directory, boundary))
    add_sheet(tables, errors, "Scenario", "Генератор сценариев", lambda: pd.DataFrame({"name": ["scenario name"], "beginDate": [boundary], "endDate": [end_boundary], "firstSummerMonth": [5], "firstWinterMonth": [10], "pumpingRate": [25]}))
    for sheet in ("Scenario_lineTankOutputs", "LineRestartRule", "LineTankOutput_lines", "emptyChangeover", "packagingChangeover"):
        add_sheet(tables, errors, sheet, "Генератор сценариев", lambda sheet=sheet: pd.DataFrame(columns=SCHEMAS[sheet]))
    if "MaterialGroup" in tables:
        add_sheet(
            tables,
            errors,
            "DemandRecord",
            name("demand"),
            lambda: make_demand(
                source_directory,
                boundary,
                end_boundary,
                tables["MaterialGroup"]["name"],
            ),
        )
    else:
        errors.append(SheetGenerationError("DemandRecord", name("demand"), "не сформирован зависимый лист «MaterialGroup»"))
    if materials is not None:
        add_sheet(tables, errors, "MaterialArrivalAssignment", name("arrival"), lambda: make_arrivals(source_directory, materials, month_start))
    else:
        errors.append(SheetGenerationError("MaterialArrivalAssignment", name("arrival"), "не сформирован зависимый лист «Material»"))
    add_sheet(tables, errors, "UnavailabilityAssignment", name("unavailability"), lambda: make_unavailability(source_directory, boundary, end_boundary))
    if "MaterialGroup" in tables:
        try:
            inputs, alternatives = make_recipe_tables(
                source_directory, tables["MaterialGroup"]["name"]
            )
        except Exception as error:  # noqa: BLE001
            errors.extend([SheetGenerationError("InputRecord", name("recipes"), str(error)), SheetGenerationError("AltInput", name("recipes"), str(error))])
        else:
            add_sheet(tables, errors, "InputRecord", name("recipes"), lambda: inputs); add_sheet(tables, errors, "AltInput", name("recipes"), lambda: alternatives)
    else:
        errors.extend([SheetGenerationError("InputRecord", name("recipes"), "не сформирован зависимый лист «MaterialGroup»"), SheetGenerationError("AltInput", name("recipes"), "не сформирован зависимый лист «MaterialGroup»")])
    if not matrix_paths(source_directory):
        errors.extend([SheetGenerationError("MaterialSwitchChangeover", "матрица переходов", "не найден обязательный файл-источник"), SheetGenerationError("ForbiddenMaterialSwitch", "матрица переходов", "не найден обязательный файл-источник")])
    else:
        try:
            switches, forbidden = make_material_switches(source_directory)
        except Exception as error:  # noqa: BLE001
            errors.extend([SheetGenerationError("MaterialSwitchChangeover", "матрица переходов", str(error)), SheetGenerationError("ForbiddenMaterialSwitch", "матрица переходов", str(error))])
        else:
            add_sheet(tables, errors, "MaterialSwitchChangeover", "матрица переходов", lambda: switches); add_sheet(tables, errors, "ForbiddenMaterialSwitch", "матрица переходов", lambda: forbidden)
    if errors:
        raise ScenarioGenerationError(errors)
    return {sheet: tables[sheet] for sheet in SHEET_ORDER}


def make_tanks(source_directory: Path) -> pd.DataFrame:
    candidates = sorted(path for path in source_directory.rglob("*.xls*") if "параметры мбх" in path.name.casefold())
    if len(candidates) != 1:
        raise ValueError("должен быть ровно один файл с маркером «параметры мбх»")
    with pd.ExcelFile(candidates[0]) as workbook:
        if not workbook.sheet_names:
            raise ValueError("файл параметров МБХ не содержит листов")
        raw = workbook.parse(sheet_name=workbook.sheet_names[-1], header=None)
    header_row = first_marker_row(raw, 0, "№ ёмкости", "Tank")
    data = raw.iloc[header_row + 1 :].copy(); data.columns = raw.iloc[header_row]
    columns = ["№ ёмкости", "Кол-во заполнения при плотности 0,898 г/смᶟ (NAOH плотности 1,45 г/смᶟ)"]
    data = normalize_excel_headers(data, columns, "Tank"); require_columns(data, columns, "Tank")
    result = pd.DataFrame({"name": "Бак " + data["№ ёмкости"].astype("string").str.replace("DEMO/", " ", regex=False).str.strip(), "capacity": (pd.to_numeric(data[columns[1]], errors="coerce") / 1_000).round(5)}).dropna(subset=["name", "capacity"])
    phantom = pd.DataFrame({"name": ["MAS E6", "Бак Лецитин", "MAS E7", "Бак 257_0", "Бак 258_0", "Вода", "Дубликат 1_Бак F 2", "Дубликат 1_Бак F 2/1", "Дубликат 2_Бак F 2", "Дубликат 2_Бак F 2/1"], "capacity": [1_000_000, 1_000_000, 1_000_000, 1_000_000, 1_000_000, 10_000_000, 4_000_000, 4_000_000, 4_000_000, 4_000_000]})
    return pd.concat([result, phantom], ignore_index=True)


def write_scenario(tables: dict[str, pd.DataFrame], output_path: Path = OUTPUT_PATH) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix="scenario_epi_", suffix=".xlsx", dir=output_path.parent, delete=False) as temp_file:
        temporary_path = Path(temp_file.name)
    try:
        with pd.ExcelWriter(temporary_path, engine="openpyxl", date_format="DD.MM.YYYY", datetime_format="DD.MM.YYYY HH:MM:SS") as writer:
            for sheet, data in tables.items(): data.to_excel(writer, sheet_name=sheet, index=False)
        os.replace(temporary_path, output_path)
    except Exception:
        temporary_path.unlink(missing_ok=True); raise


def create_scenario(source_directory: Path | None = None, output_path: Path | None = None, today: pd.Timestamp | None = None) -> Path:
    path = Path(output_path) if output_path is not None else OUTPUT_PATH
    if today is None:
        tables = build_tables(source_directory=source_directory)
    else:
        tables = build_tables(today=today, source_directory=source_directory)
    write_scenario(tables, path)
    return path


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Генератор файла сценария ЭПИ")
    parser.add_argument("--source-dir", type=Path); parser.add_argument("--output-path", type=Path)
    parser.add_argument("--as-of-date", type=pd.Timestamp, help="Дата запуска для выбора следующего месяца планирования (ГГГГ-ММ-ДД).")
    arguments = parser.parse_args()
    print(f"Создан файл: {create_scenario(arguments.source_dir, arguments.output_path, arguments.as_of_date)}")
