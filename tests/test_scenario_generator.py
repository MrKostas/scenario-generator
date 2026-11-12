"""Проверки публичного API генератора сценария."""

from __future__ import annotations

import importlib.util
import sys
import types
import unittest
from pathlib import Path
from unittest.mock import patch

PROJECT_DIRECTORY = Path(__file__).resolve().parents[1]
GENERATOR_PATH = PROJECT_DIRECTORY / "scenario-generator.py"


def load_generator_module():
    specification = importlib.util.spec_from_file_location(
        "scenario_generator_test_module", GENERATOR_PATH
    )
    if specification is None or specification.loader is None:
        raise ImportError(f"Не удалось загрузить генератор: {GENERATOR_PATH}")
    module = importlib.util.module_from_spec(specification)
    sys.modules[specification.name] = module
    pandas_stub = types.ModuleType("pandas")
    with patch.dict(sys.modules, {"pandas": pandas_stub}):
        specification.loader.exec_module(module)
    return module


class CreateScenarioTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.generator = load_generator_module()

    def test_create_scenario_passes_source_and_output_paths(self) -> None:
        source_directory = Path("/tmp/source")
        output_path = Path("/tmp/source/Real Scenario Евдаково.xlsx")
        tables = object()

        with (
            patch.object(
                self.generator, "build_tables", return_value=tables
            ) as build_tables,
            patch.object(self.generator, "write_scenario") as write_scenario,
        ):
            created_path = self.generator.create_scenario(
                source_directory=source_directory,
                output_path=output_path,
            )

        build_tables.assert_called_once_with(source_directory=source_directory)
        write_scenario.assert_called_once_with(tables, output_path)
        self.assertEqual(created_path, output_path)

    def test_accumulated_errors_keep_existing_output_unchanged(self) -> None:
        output_path = Path("/tmp/Real Scenario Евдаково.xlsx")
        errors = [
            self.generator.SheetGenerationError(
                sheet="Material",
                source="УС-САП-Информация о продуктах.xlsx",
                reason="отсутствует столбец shelfLife",
            ),
            self.generator.SheetGenerationError(
                sheet="Tank",
                source="УС-САП-Трубы и баки.xlsx",
                reason="не найден лист баки",
            ),
        ]
        generation_error = self.generator.ScenarioGenerationError(errors)

        with (
            patch.object(
                self.generator, "build_tables", side_effect=generation_error
            ),
            patch.object(self.generator, "write_scenario") as write_scenario,
        ):
            with self.assertRaises(self.generator.ScenarioGenerationError) as raised:
                self.generator.create_scenario(output_path=output_path)

        write_scenario.assert_not_called()
        message = str(raised.exception)
        self.assertIn("листа «Material»", message)
        self.assertIn("УС-САП-Информация о продуктах.xlsx", message)
        self.assertIn("листа «Tank»", message)
        self.assertIn("УС-САП-Трубы и баки.xlsx", message)


if __name__ == "__main__":
    unittest.main()
