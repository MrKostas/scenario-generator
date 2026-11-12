"""Проверки отображения накопленных ошибок tolerant ETL в Desktop UI."""

from __future__ import annotations

import queue
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import desktop_app


class BatchGenerationError(Exception):
    """Минимальная модель исключения ETL с накопленными ошибками листов."""

    def __init__(self) -> None:
        self.errors = [object(), object()]
        super().__init__(
            "При генерации листа «Material» произошли ошибки. "
            "Исправьте соответствующий файл-источник «products.xlsx». "
            "Причина: missing shelfLife\n\n"
            "При генерации листа «Tank» произошли ошибки. "
            "Исправьте соответствующий файл-источник «tanks.xlsx». "
            "Причина: missing capacity"
        )


def make_app_without_tk() -> desktop_app.ScenarioGeneratorApp:
    """Создаёт объект для проверки handlers без окна Tk/display-сервера."""
    app = object.__new__(desktop_app.ScenarioGeneratorApp)
    app.root = object()
    app.processing_started_at = None
    return app


class TolerantDesktopTests(unittest.TestCase):
    def test_success_with_warnings_shows_one_warning_and_keeps_result_path(self):
        app = make_app_without_tk()
        output = Path("result.xlsx")
        with (
            patch.object(app, "_finish_generation"),
            patch.object(app, "_set_status") as status,
            patch.object(desktop_app, "reveal_output_file") as reveal,
            patch.object(desktop_app.messagebox, "showinfo") as info,
            patch.object(desktop_app.messagebox, "showwarning") as warning,
        ):
            app._handle_success(output, warnings=("Сумма 0.98", "Нет операции"))
        info.assert_not_called()
        warning.assert_called_once()
        self.assertEqual(warning.call_args.args[0], "Файл создан с предупреждениями")
        self.assertIn(str(output), warning.call_args.args[1])
        self.assertIn("Сумма 0.98", warning.call_args.args[1])
        self.assertIn("Нет операции", warning.call_args.args[1])
        self.assertIn("предупреждениями", status.call_args.args[0])
        reveal.assert_called_once_with(output)

    def test_worker_collects_evdakovo_warnings_and_keeps_epi_api(self):
        for site in ("Евдаково", "ЭПИ"):
            with self.subTest(site=site):
                app = make_app_without_tk()
                app.result_queue = queue.Queue()
                calls = []

                def create_scenario(_calls=calls, **kwargs):
                    _calls.append(kwargs)
                    if "on_warning" in kwargs:
                        kwargs["on_warning"]("Сумма 0.98")
                    return kwargs["output_path"]

                app.generator_modules = {
                    site: SimpleNamespace(create_scenario=create_scenario)
                }
                app._run_generation(Path("sources"), site)
                event, (output, warnings) = app.result_queue.get_nowait()
                self.assertEqual(event, "success")
                self.assertEqual(
                    output,
                    Path("sources")
                    / desktop_app.SITE_CONFIGURATIONS[site].output_filename,
                )
                self.assertEqual(
                    warnings, ("Сумма 0.98",) if site == "Евдаково" else ()
                )
                self.assertEqual("on_warning" in calls[0], site == "Евдаково")

    def test_poll_passes_collected_warnings_to_success_handler(self):
        app = make_app_without_tk()
        app.result_queue = queue.Queue()
        app.result_queue.put(("success", (Path("result.xlsx"), ("Сумма 0.98",))))
        with patch.object(app, "_handle_success") as success:
            app._poll_generation_result()
        success.assert_called_once_with(Path("result.xlsx"), warnings=("Сумма 0.98",))

    def test_batch_failure_shows_exactly_one_warning_with_all_sources(self) -> None:
        app = make_app_without_tk()
        error = BatchGenerationError()

        with (
            patch.object(app, "_finish_generation") as finish_generation,
            patch.object(app, "_set_status") as set_status,
            patch.object(desktop_app.messagebox, "showwarning") as showwarning,
            patch.object(desktop_app.messagebox, "showerror") as showerror,
        ):
            app._handle_failure(error)

        finish_generation.assert_called_once_with()
        set_status.assert_called_once_with(
            "Файл сценария не создан: исправьте ошибки в источниках и повторите.",
            desktop_app.ERROR,
        )
        showwarning.assert_called_once_with(
            "Файл сценария не создан", str(error), parent=app.root
        )
        showerror.assert_not_called()
        warning_text = showwarning.call_args.args[1]
        self.assertIn("«Material»", warning_text)
        self.assertIn("«products.xlsx»", warning_text)
        self.assertIn("«Tank»", warning_text)
        self.assertIn("«tanks.xlsx»", warning_text)

    def test_success_shows_no_warning_and_reports_saved_path(self) -> None:
        app = make_app_without_tk()
        output_path = Path("/tmp/Real Scenario Евдаково.xlsx")

        with (
            patch.object(app, "_finish_generation") as finish_generation,
            patch.object(app, "_set_status") as set_status,
            patch.object(desktop_app, "reveal_output_file") as reveal_output_file,
            patch.object(desktop_app.messagebox, "showinfo") as showinfo,
            patch.object(desktop_app.messagebox, "showwarning") as showwarning,
        ):
            app._handle_success(output_path)

        finish_generation.assert_called_once_with()
        set_status.assert_called_once_with(
            f"Обработка завершена. Файл сохранён:\n{output_path}", desktop_app.PRIMARY
        )
        reveal_output_file.assert_called_once_with(output_path)
        showinfo.assert_called_once_with(
            "Обработка завершена",
            f"Файл сценария сохранён по пути:\n{output_path}",
            parent=app.root,
        )
        showwarning.assert_not_called()


if __name__ == "__main__":
    unittest.main()
