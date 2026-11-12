"""Проверки вспомогательной логики Desktop-интерфейса."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import desktop_app


class ProcessingEstimateTests(unittest.TestCase):
    def test_estimate_uses_excel_source_size(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            source_directory = Path(temporary_directory)
            (source_directory / "source.xlsx").write_bytes(b"0" * 3 * 1024 * 1024)

            estimate = desktop_app.estimate_processing_seconds(source_directory)

        self.assertEqual(estimate, 8)

    def test_previous_duration_has_priority_over_source_size(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            estimate = desktop_app.estimate_processing_seconds(
                Path(temporary_directory),
                previous_duration_seconds=100,
            )

        self.assertEqual(estimate, 111)

    def test_short_previous_duration_is_not_rounded_to_a_minute(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            estimate = desktop_app.estimate_processing_seconds(
                Path(temporary_directory),
                previous_duration_seconds=1,
            )

        self.assertEqual(estimate, 5)

    def test_duration_is_formatted_for_status(self) -> None:
        self.assertEqual(desktop_app.format_duration(125), "2 мин 05 сек")


class SiteConfigurationTests(unittest.TestCase):
    def test_site_configurations_use_separate_output_and_pipeline_paths(self) -> None:
        evdakovo = desktop_app.SITE_CONFIGURATIONS["Евдаково"]
        epi = desktop_app.SITE_CONFIGURATIONS["ЭПИ"]

        self.assertEqual(evdakovo.output_filename, "Real Scenario Евдаково.xlsx")
        self.assertEqual(epi.output_filename, "Real Scenario ЭПИ.xlsx")
        self.assertNotEqual(evdakovo.generator_path, epi.generator_path)


class FileManagerCommandTests(unittest.TestCase):
    def test_macos_reveals_and_selects_result_file(self) -> None:
        output_path = Path("/tmp/Real Scenario.xlsx")

        command = desktop_app.file_manager_command(output_path, platform="darwin")

        self.assertEqual(command, ["open", "-R", str(output_path)])

    def test_windows_reveals_and_selects_result_file(self) -> None:
        output_path = Path("C:/scenario/Real Scenario.xlsx")

        command = desktop_app.file_manager_command(output_path, platform="win32")

        self.assertEqual(command, ["explorer", "/select,", str(output_path)])

    def test_linux_opens_result_directory_when_selection_is_unavailable(self) -> None:
        output_path = Path("/tmp/Real Scenario.xlsx")

        command = desktop_app.file_manager_command(output_path, platform="linux")

        self.assertEqual(command, ["xdg-open", str(output_path.parent)])


class GenerationFailureDialogTests(unittest.TestCase):
    def test_sheet_errors_use_single_warning_dialog(self) -> None:
        app = object.__new__(desktop_app.ScenarioGeneratorApp)
        app.root = object()
        app._finish_generation = Mock()
        app._set_status = Mock()
        error = RuntimeError("all sheet errors")
        error.errors = [object(), object()]

        with patch.object(desktop_app.messagebox, "showwarning") as showwarning:
            desktop_app.ScenarioGeneratorApp._handle_failure(app, error)

        app._finish_generation.assert_called_once_with()
        app._set_status.assert_called_once_with(
            "Файл сценария не создан: исправьте ошибки в источниках и повторите.",
            desktop_app.ERROR,
        )
        showwarning.assert_called_once_with(
            "Файл сценария не создан", str(error), parent=app.root
        )


if __name__ == "__main__":
    unittest.main()
