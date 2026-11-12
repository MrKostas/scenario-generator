"""Контрактные проверки остановки ETL при schema drift источников."""

from __future__ import annotations

import importlib.util
import sys
import tempfile
import types
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

PROJECT_DIRECTORY = Path(__file__).resolve().parents[1]
GENERATOR_PATH = PROJECT_DIRECTORY / "scenario-generator.py"


def load_generator_module():
    specification = importlib.util.spec_from_file_location(
        "scenario_generator_schema_test_module", GENERATOR_PATH
    )
    if specification is None or specification.loader is None:
        raise ImportError(f"Не удалось загрузить генератор: {GENERATOR_PATH}")
    module = importlib.util.module_from_spec(specification)
    sys.modules[specification.name] = module
    pandas_stub = types.ModuleType("pandas")
    with patch.dict(sys.modules, {"pandas": pandas_stub}):
        specification.loader.exec_module(module)
    return module


class FakeDataFrame:
    """Минимальная модель DataFrame для проверки валидации заголовков."""

    def __init__(self, columns: list[str]) -> None:
        self.columns = columns

    def rename(self, *, columns: dict[str, str]):
        return FakeDataFrame([columns.get(column, column) for column in self.columns])

    def __getitem__(self, columns: list[str]):
        return FakeDataFrame(columns)


class FakePlanningDate:
    """Достаточная модель Timestamp для маршрута build_tables без pandas."""

    def normalize(self):
        return self

    def replace(self, **_kwargs: object):
        return self

    def __add__(self, _other: object):
        return self

    def date(self):
        return date(2026, 3, 1)


class FakePandasForBuild:
    def __init__(self) -> None:
        self.offsets = types.SimpleNamespace(
            MonthBegin=lambda _count: object(), MonthEnd=lambda _count: object()
        )

    @staticmethod
    def Timedelta(**_kwargs: object) -> object:
        return object()

    @staticmethod
    def DataFrame(
        data: object = None, columns: list[str] | None = None
    ) -> FakeDataFrame:
        if columns is not None:
            return FakeDataFrame(columns)
        if isinstance(data, dict):
            return FakeDataFrame(list(data))
        return FakeDataFrame([])


class SchemaFailureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.generator = load_generator_module()

    def test_missing_required_column_stops_validation(self) -> None:
        data = FakeDataFrame(["name"])

        with self.assertRaisesRegex(
            ValueError, r"Material: отсутствуют столбцы \['shelfLife'\]"
        ):
            self.generator.require_columns(data, ["name", "shelfLife"], "Material")

    def test_missing_worksheet_error_is_not_silently_skipped(self) -> None:
        with patch.object(
            self.generator,
            "read_source",
            side_effect=ValueError("Worksheet named 'modeChange' not found"),
        ):
            with self.assertRaisesRegex(
                ValueError, r"Worksheet named 'modeChange' not found"
            ):
                self.generator.read_contract_source(
                    "changeovers",
                    Path("/tmp/sources"),
                    "modeChange",
                    ["previousMode"],
                )

    def test_duplicate_normalized_headers_stop_validation(self) -> None:
        data = FakeDataFrame([" name ", "NAME"])

        with self.assertRaisesRegex(
            ValueError, r"Material: дублирующиеся заголовки после нормализации: \['name'\]"
        ):
            self.generator.normalize_excel_headers(data, ["name"], "Material")

    def test_schema_failure_preserves_existing_output(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_path = Path(temporary_directory) / "Real Scenario Евдаково.xlsx"
            output_path.write_bytes(b"previous scenario")

            with (
                patch.object(
                    self.generator,
                    "build_tables",
                    side_effect=ValueError("Material: отсутствуют столбцы ['shelfLife']"),
                ),
                patch.object(self.generator, "write_scenario") as write_scenario,
            ):
                with self.assertRaisesRegex(ValueError, r"Material: отсутствуют столбцы"):
                    self.generator.create_scenario(
                        source_directory=Path(temporary_directory),
                        output_path=output_path,
                    )

            write_scenario.assert_not_called()
            self.assertEqual(output_path.read_bytes(), b"previous scenario")


class TolerantGenerationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.generator = load_generator_module()

    def test_independent_sheet_errors_are_accumulated_and_next_sheet_runs(self) -> None:
        tables = {}
        errors = []
        attempted = []

        def failing_producer(name: str):
            def produce():
                attempted.append(name)
                raise ValueError(f"broken {name}")

            return produce

        def successful_producer():
            attempted.append("LineGroup")
            return FakeDataFrame(["name"])

        self.generator.add_sheet(
            tables, errors, "Material", "products.xlsx", failing_producer("Material")
        )
        self.generator.add_sheet(
            tables, errors, "Tank", "tanks.xlsx", failing_producer("Tank")
        )
        created = self.generator.add_sheet(
            tables, errors, "LineGroup", "generator", successful_producer
        )

        self.assertEqual(attempted, ["Material", "Tank", "LineGroup"])
        self.assertEqual([error.sheet for error in errors], ["Material", "Tank"])
        self.assertEqual([error.source for error in errors], ["products.xlsx", "tanks.xlsx"])
        self.assertIsNotNone(created)
        self.assertIn("LineGroup", tables)

    def test_accumulated_error_message_lists_each_sheet_and_source(self) -> None:
        error = self.generator.ScenarioGenerationError(
            [
                self.generator.SheetGenerationError("Material", "products.xlsx", "missing name"),
                self.generator.SheetGenerationError("Tank", "tanks.xlsx", "missing capacity"),
            ]
        )

        self.assertEqual(len(error.errors), 2)
        self.assertIn("листа «Material»", str(error))
        self.assertIn("источник «products.xlsx»", str(error))
        self.assertIn("листа «Tank»", str(error))
        self.assertIn("источник «tanks.xlsx»", str(error))

    def test_accumulated_error_does_not_replace_existing_output(self) -> None:
        errors = [
            self.generator.SheetGenerationError(
                "Material", "products.xlsx", "missing shelfLife"
            )
        ]
        with tempfile.TemporaryDirectory() as temporary_directory:
            output_path = Path(temporary_directory) / "Real Scenario Евдаково.xlsx"
            output_path.write_bytes(b"previous scenario")

            with (
                patch.object(
                    self.generator,
                    "build_tables",
                    side_effect=self.generator.ScenarioGenerationError(errors),
                ),
                patch.object(self.generator, "write_scenario") as write_scenario,
            ):
                with self.assertRaises(self.generator.ScenarioGenerationError):
                    self.generator.create_scenario(
                        source_directory=Path(temporary_directory),
                        output_path=output_path,
                    )

            write_scenario.assert_not_called()
            self.assertEqual(output_path.read_bytes(), b"previous scenario")

    def test_missing_material_switch_matrix_blocks_export_and_preserves_output(self) -> None:
        """Регрессия E2E: нет ЦПСЖ-матриц — нет частичной книги сценария."""
        with tempfile.TemporaryDirectory() as temporary_directory:
            source_directory = Path(temporary_directory)
            output_path = source_directory / "Real Scenario Евдаково.xlsx"
            output_path.write_bytes(b"previous scenario")

            with (
                patch.object(self.generator, "pd", FakePandasForBuild()),
                patch.object(self.generator, "add_sheet", return_value=None),
            ):
                with self.assertRaises(self.generator.ScenarioGenerationError) as raised:
                    self.generator.build_tables(
                        today=FakePlanningDate(), source_directory=source_directory
                    )

            matrix_errors = [
                error
                for error in raised.exception.errors
                if error.sheet == "MaterialSwitchChangeover"
            ]
            self.assertEqual(
                [(error.sheet, error.source) for error in matrix_errors],
                [("MaterialSwitchChangeover", "матрица переходов ЦПСЖ")],
            )
            self.assertTrue(
                all(
                    "не найден обязательный файл-источник" in error.reason
                    for error in matrix_errors
                )
            )

            with (
                patch.object(
                    self.generator, "build_tables", side_effect=raised.exception
                ),
                patch.object(self.generator, "write_scenario") as write_scenario,
            ):
                with self.assertRaises(self.generator.ScenarioGenerationError):
                    self.generator.create_scenario(
                        source_directory=source_directory, output_path=output_path
                    )

            write_scenario.assert_not_called()
            self.assertEqual(output_path.read_bytes(), b"previous scenario")


if __name__ == "__main__":
    unittest.main()
