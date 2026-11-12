"""Desktop-интерфейс для формирования файла сценария."""

from __future__ import annotations

import argparse
import importlib.util
import math
import queue
import subprocess
import sys
import threading
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from types import ModuleType

BUNDLE_DIRECTORY = Path(__file__).resolve().parent
APPLICATION_DIRECTORY = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else BUNDLE_DIRECTORY
)
GENERATOR_PATH = BUNDLE_DIRECTORY / "scenario-generator.py"
EPI_GENERATOR_PATH = BUNDLE_DIRECTORY / "scenario-generator-epi.py"

CARD_BACKGROUND = "#FFFFFF"
STATUS_BACKGROUND = "#F1F2F1"
PRIMARY = "#588157"
PRIMARY_ACTIVE = "#3A5A40"
PRIMARY_TEXT = "#FFFFFF"
TEXT = "#344E41"
MUTED_TEXT = "#3A5A40"
DIVIDER = "#A3B18A"
OUTLINE = "#344E41"
DISABLED_TEXT = "#78837D"
ERROR = "#B42318"
FONT_FAMILY = "Segoe UI"
HEADING_FONT = (FONT_FAMILY, 13, "bold")
BODY_FONT = (FONT_FAMILY, 11)
BUTTON_FONT = (FONT_FAMILY, 11, "bold")
EXCEL_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}
WINDOW_WIDTH = 980
WINDOW_HEIGHT = 540
CONTENT_HORIZONTAL_INSET = 54
CONTENT_VERTICAL_INSET = 24
CONTENT_WIDTH = WINDOW_WIDTH - 2 * CONTENT_HORIZONTAL_INSET
CONTENT_HEIGHT = WINDOW_HEIGHT - 2 * CONTENT_VERTICAL_INSET


@dataclass(frozen=True)
class SiteConfiguration:
    """Параметры запуска ETL-пайплайна для конкретной площадки."""

    source_directory: Path
    output_filename: str
    generator_path: Path


SITE_CONFIGURATIONS = {
    "Евдаково": SiteConfiguration(
        source_directory=APPLICATION_DIRECTORY / "data" / "raw" / "Евдаково",
        output_filename="Real Scenario Евдаково.xlsx",
        generator_path=GENERATOR_PATH,
    ),
    "ЭПИ": SiteConfiguration(
        source_directory=APPLICATION_DIRECTORY / "data" / "raw" / "ЭПИ",
        output_filename="Real Scenario ЭПИ.xlsx",
        generator_path=EPI_GENERATOR_PATH,
    ),
}


def current_time() -> datetime:
    """Возвращает локальное timezone-aware время для статуса обработки."""
    return datetime.now(timezone.utc).astimezone()


def estimate_processing_seconds(
    source_directory: Path,
    previous_duration_seconds: float | None = None,
) -> int:
    """Возвращает консервативную оценку времени обработки в секундах."""
    if previous_duration_seconds is not None:
        return max(5, min(600, math.ceil(previous_duration_seconds * 1.1)))

    total_size = sum(
        path.stat().st_size
        for path in source_directory.rglob("*")
        if path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS
    )
    size_in_megabytes = total_size / (1024 * 1024)
    return max(5, min(60, 5 + math.ceil(size_in_megabytes)))


def format_duration(seconds: int) -> str:
    """Представляет небольшую длительность в форме, удобной для статуса."""
    minutes, remainder = divmod(max(seconds, 0), 60)
    if minutes:
        return f"{minutes} мин {remainder:02d} сек"
    return f"{remainder} сек"


def file_manager_command(
    output_path: Path,
    platform: str | None = None,
) -> list[str] | None:
    """Возвращает команду для показа файла в файловом менеджере ОС."""
    operating_system = platform or sys.platform
    path = str(output_path)
    if operating_system == "darwin":
        return ["open", "-R", path]
    if operating_system.startswith("win"):
        return ["explorer", "/select,", path]
    if operating_system.startswith("linux"):
        return ["xdg-open", str(output_path.parent)]
    return None


def reveal_output_file(output_path: Path) -> None:
    """Открывает папку результата; выделение файла поддерживается не всеми ОС."""
    command = file_manager_command(output_path)
    if command is None:
        return
    try:
        subprocess.Popen(command)
    except OSError:
        # Открытие проводника не влияет на уже успешный ETL-результат.
        return


def load_generator_module(generator_path: Path = GENERATOR_PATH) -> ModuleType:
    """Загружает ETL-модуль, имя файла которого содержит дефис."""
    if not generator_path.is_file():
        raise FileNotFoundError(
            "Для выбранной площадки ещё не подключён отдельный ETL-пайплайн: "
            f"{generator_path.name}"
        )

    module_name = f"scenario_generator_{generator_path.stem.replace('-', '_')}"
    specification = importlib.util.spec_from_file_location(module_name, generator_path)
    if specification is None or specification.loader is None:
        raise ImportError(f"Не удалось загрузить генератор: {generator_path}")

    module = importlib.util.module_from_spec(specification)
    sys.modules[specification.name] = module
    specification.loader.exec_module(module)
    return module


class ScenarioGeneratorApp:
    """Окно запуска ETL-пайплайна в отдельном потоке."""

    def __init__(self, root: tk.Tk, generator_module: ModuleType) -> None:
        self.root = root
        self.generator_modules = {"Евдаково": generator_module}
        self.selected_site = tk.StringVar(value="Евдаково")
        self.status_text = tk.StringVar(value="")
        self.last_duration_seconds: float | None = None
        self.processing_started_at: datetime | None = None
        self.expected_finish_at: datetime | None = None
        self.status_timer_job: str | None = None
        self.result_poll_job: str | None = None
        self.result_queue: queue.Queue[
            tuple[str, tuple[Path, tuple[str, ...]] | Exception]
        ] = queue.Queue()

        self.root.title("Генератор сценариев")
        self.root.configure(background=CARD_BACKGROUND)
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.resizable(False, False)

        self._build_interface()

    def _build_interface(self) -> None:
        self.content = tk.Frame(self.root, background=CARD_BACKGROUND)
        self.content.place(
            x=CONTENT_HORIZONTAL_INSET,
            y=CONTENT_VERTICAL_INSET,
            width=CONTENT_WIDTH,
            height=CONTENT_HEIGHT,
        )
        self.site_heading = self._create_heading("Выберите площадку")
        self.site_heading.place(x=0, y=53, width=270, height=29)
        self.scenario_heading = self._create_heading("Формирование файла-сценария")
        self.scenario_heading.place(x=310, y=53, width=538, height=29)

        self._configure_site_selector()
        self.site_selector = ttk.Combobox(
            self.content,
            textvariable=self.selected_site,
            values=tuple(SITE_CONFIGURATIONS),
            state="readonly",
            style="Site.TCombobox",
            font=BODY_FONT,
        )
        self.site_selector.place(x=0, y=98, width=270, height=44)
        self.site_selector.bind("<<ComboboxSelected>>", self._on_site_selected)

        self.scenario_text = tk.StringVar(
            value=("Выберите папку с Excel-формами, чтобы создать книгу сценария.")
        )
        self.scenario_label = self._create_text_label(self.scenario_text)
        self.scenario_label.place(x=310, y=98, width=250, height=106)
        self.generate_button = tk.Button(
            self.content,
            text="Сформировать сценарий",
            command=self.select_source_directory,
            background=PRIMARY,
            activebackground=PRIMARY_ACTIVE,
            font=BUTTON_FONT,
            foreground=PRIMARY_TEXT,
            activeforeground=PRIMARY_TEXT,
            disabledforeground=DISABLED_TEXT,
            relief="flat",
            borderwidth=0,
            cursor="hand2",
        )
        self.generate_button.place(x=604, y=98, width=244, height=64)

        self.site_copy_one = self._create_text_label(
            "У каждой площадки свои Excel-формы и правила обработки."
        )
        self.site_copy_one.place(x=0, y=204, width=270, height=86)
        self.report_heading = self._create_heading(
            "Формирование динамической отчетности"
        )
        self.report_heading.place(x=310, y=261, width=538, height=29)

        self.site_copy_two = self._create_text_label(
            "Приложение сохранит книгу в выбранной папке."
        )
        self.site_copy_two.place(x=0, y=306, width=270, height=66)
        self.report_copy = self._create_text_label(
            "Динамический отчёт пока недоступен."
        )
        self.report_copy.place(x=310, y=306, width=250, height=66)
        self.report_button = tk.Button(
            self.content,
            text="Сформировать отчет",
            command=self._show_report_unavailable,
            background=CARD_BACKGROUND,
            activebackground="#EEF1EC",
            font=BUTTON_FONT,
            foreground=TEXT,
            activeforeground=TEXT,
            relief="solid",
            borderwidth=1,
            cursor="hand2",
        )
        self.report_button.place(x=604, y=306, width=244, height=64)

        self.vertical_divider = tk.Frame(self.content, background=DIVIDER, width=1)
        self.right_divider = tk.Frame(self.content, background=DIVIDER, height=1)

        self.status_panel = tk.Frame(self.content, background=STATUS_BACKGROUND)
        self.status_panel.columnconfigure(1, weight=1)
        self.loading_indicator = ttk.Progressbar(
            self.status_panel,
            mode="indeterminate",
            length=64,
        )
        self.loading_indicator.grid(row=0, column=0, padx=(12, 0), pady=8)
        self.loading_indicator.grid_remove()
        self.status_label = tk.Label(
            self.status_panel,
            textvariable=self.status_text,
            background=STATUS_BACKGROUND,
            foreground=MUTED_TEXT,
            justify="left",
            anchor="w",
            font=BODY_FONT,
            padx=16,
            pady=8,
        )
        self.status_label.grid(row=0, column=1, sticky="ew")

        self.scenario_label.configure(wraplength=238)
        self.site_copy_one.configure(wraplength=260)
        self.site_copy_two.configure(wraplength=260)
        self.report_copy.configure(wraplength=238)
        self.status_label.configure(wraplength=790)
        self.vertical_divider.place(x=290, y=0, width=1, height=CONTENT_HEIGHT)
        self.right_divider.place(x=310, y=204, width=CONTENT_WIDTH - 310, height=1)
        self._set_status(
            "Готово к запуску. Выберите папку с Excel-источниками.",
            MUTED_TEXT,
        )

    def _create_heading(self, text: str) -> tk.Label:
        return tk.Label(
            self.content,
            text=text,
            background=CARD_BACKGROUND,
            foreground=MUTED_TEXT,
            font=HEADING_FONT,
            anchor="w",
        )

    def _create_text_label(
        self,
        text: str | tk.StringVar,
    ) -> tk.Label:
        return tk.Label(
            self.content,
            text=text if isinstance(text, str) else "",
            textvariable=text if isinstance(text, tk.StringVar) else None,
            background=CARD_BACKGROUND,
            foreground=MUTED_TEXT,
            font=BODY_FONT,
            justify="left",
            anchor="nw",
        )

    def _configure_site_selector(self) -> None:
        style = ttk.Style(self.root)
        style.configure(
            "Site.TCombobox",
            fieldbackground=CARD_BACKGROUND,
            background=CARD_BACKGROUND,
            foreground=TEXT,
            bordercolor=OUTLINE,
            lightcolor=OUTLINE,
            darkcolor=OUTLINE,
            arrowcolor=TEXT,
            padding=(14, 11),
        )
        style.map(
            "Site.TCombobox",
            fieldbackground=[("readonly", CARD_BACKGROUND)],
            foreground=[("readonly", TEXT)],
            selectbackground=[("readonly", CARD_BACKGROUND)],
            selectforeground=[("readonly", TEXT)],
        )

    def toggle_site(self) -> None:
        """Переключает площадку и тем самым будущий ETL-пайплайн."""
        site_names = tuple(SITE_CONFIGURATIONS)
        current_index = site_names.index(self.selected_site.get())
        next_site = site_names[(current_index + 1) % len(site_names)]
        self.selected_site.set(next_site)
        self._on_site_selected()

    def _on_site_selected(
        self,
        _event: tk.Event[tk.Misc] | None = None,
    ) -> None:
        """Обновляет пояснение после смены площадки в выпадающем списке."""
        self.scenario_text.set(
            "Файл-сценарий формируется из исходных форм СЭДО.\n\n"
            "Нажмите кнопку и выберите папку с исходными данными."
        )
        self._set_status(
            "Выбрана площадка «"
            f"{self.selected_site.get()}». Выберите папку с Excel-источниками.",
            MUTED_TEXT,
        )
        self.status_panel.place_forget()

    def _current_site_configuration(self) -> SiteConfiguration:
        return SITE_CONFIGURATIONS[self.selected_site.get()]

    def select_source_directory(self) -> None:
        configuration = self._current_site_configuration()
        initial_directory = (
            configuration.source_directory
            if configuration.source_directory.is_dir()
            else APPLICATION_DIRECTORY
        )
        source_directory = filedialog.askdirectory(
            parent=self.root,
            title="Выберите папку с Excel-источниками",
            initialdir=initial_directory,
            mustexist=True,
        )
        if not source_directory:
            self._set_status("Выбор папки отменён.", MUTED_TEXT)
            self._show_status_panel()
            return

        selected_path = Path(source_directory)
        self.scenario_text.set(f"Источник: {selected_path}")
        self._start_generation(selected_path, self.selected_site.get())

    def _start_generation(self, source_directory: Path, site_name: str) -> None:
        self.generate_button.configure(state="disabled")
        self.site_selector.configure(state="disabled")
        estimate = estimate_processing_seconds(
            source_directory,
            self.last_duration_seconds,
        )
        self.processing_started_at = current_time()
        self.expected_finish_at = self.processing_started_at + timedelta(
            seconds=estimate
        )
        self._show_status_panel()
        self.loading_indicator.grid()
        self.loading_indicator.start(12)
        self._update_processing_status()
        self._poll_generation_result()

        worker = threading.Thread(
            target=self._run_generation,
            args=(source_directory, site_name),
            daemon=True,
        )
        worker.start()

    def _run_generation(self, source_directory: Path, site_name: str) -> None:
        configuration = SITE_CONFIGURATIONS[site_name]
        output_path = source_directory / configuration.output_filename
        warnings: list[str] = []
        try:
            generator_module = self.generator_modules.get(site_name)
            if generator_module is None:
                generator_module = load_generator_module(configuration.generator_path)
                self.generator_modules[site_name] = generator_module
            if site_name == "Евдаково":
                created_path = generator_module.create_scenario(
                    source_directory=source_directory,
                    output_path=output_path,
                    on_warning=warnings.append,
                )
            else:
                created_path = generator_module.create_scenario(
                    source_directory=source_directory,
                    output_path=output_path,
                )
        except Exception as error:  # noqa: BLE001
            self.result_queue.put(("failure", error))
            return
        self.result_queue.put(
            ("success", (Path(created_path), tuple(dict.fromkeys(warnings))))
        )

    def _poll_generation_result(self) -> None:
        try:
            result_type, result = self.result_queue.get_nowait()
        except queue.Empty:
            self.result_poll_job = self.root.after(100, self._poll_generation_result)
            return

        self.result_poll_job = None
        if result_type == "success":
            output_path, warnings = result
            self._handle_success(output_path, warnings=warnings)
        else:
            self._handle_failure(result)

    def _handle_success(
        self, output_path: Path, warnings: tuple[str, ...] = ()
    ) -> None:
        if self.processing_started_at is not None:
            self.last_duration_seconds = (
                current_time() - self.processing_started_at
            ).total_seconds()
        self._finish_generation()
        status = "Файл создан с предупреждениями" if warnings else "Обработка завершена"
        self._set_status(f"{status}. Файл сохранён:\n{output_path}", PRIMARY)
        reveal_output_file(output_path)
        if warnings:
            messagebox.showwarning(
                "Файл создан с предупреждениями",
                f"Файл сценария сохранён по пути:\n{output_path}\n\n"
                + "\n\n".join(warnings),
                parent=self.root,
            )
            return
        messagebox.showinfo(
            "Обработка завершена",
            f"Файл сценария сохранён по пути:\n{output_path}",
            parent=self.root,
        )

    def _handle_failure(self, error: Exception) -> None:
        self._finish_generation()
        sheet_errors = getattr(error, "errors", None)
        if sheet_errors:
            self._set_status(
                "Файл сценария не создан: исправьте ошибки в источниках и повторите.",
                ERROR,
            )
            messagebox.showwarning(
                "Файл сценария не создан",
                str(error),
                parent=self.root,
            )
            return
        self._set_status(
            "Не удалось сформировать данные. Проверьте источники и повторите.",
            ERROR,
        )
        messagebox.showerror(
            "Не удалось сформировать данные",
            "Проверьте выбранную папку и обязательные Excel-источники."
            f"\n\nПричина: {error}",
            parent=self.root,
        )

    def _finish_generation(self) -> None:
        if self.status_timer_job is not None:
            self.root.after_cancel(self.status_timer_job)
            self.status_timer_job = None
        if self.result_poll_job is not None:
            self.root.after_cancel(self.result_poll_job)
            self.result_poll_job = None
        self.loading_indicator.stop()
        self.loading_indicator.grid_remove()
        self.generate_button.configure(state="normal")
        self.site_selector.configure(state="readonly")

    def _update_processing_status(self) -> None:
        if self.expected_finish_at is None:
            return

        remaining_seconds = math.ceil(
            (self.expected_finish_at - current_time()).total_seconds()
        )
        if remaining_seconds > 0:
            message = (
                "Формируем данные…\n"
                f"Ориентировочное завершение: {self.expected_finish_at:%H:%M} "
                f"· осталось около {format_duration(remaining_seconds)}"
            )
        else:
            message = (
                "Формирование продолжается дольше первоначальной оценки…\n"
                "Пожалуйста, дождитесь сообщения о завершении."
            )
        self._set_status(message, TEXT)
        self.status_timer_job = self.root.after(1000, self._update_processing_status)

    def _set_status(self, message: str, color: str) -> None:
        self.status_text.set(message)
        self.status_label.configure(foreground=color)

    def _show_status_panel(self) -> None:
        self.status_panel.place(x=0, y=388, width=CONTENT_WIDTH, height=80)

    def _show_report_unavailable(self) -> None:
        """Объясняет, что второй pipeline пока не подключён к заготовке."""
        self._set_status("Формирование отчёта пока не подключено.", MUTED_TEXT)
        self._show_status_panel()
        messagebox.showinfo(
            "Формирование отчёта",
            "Pipeline для формирования динамического отчёта пока не подключён.",
            parent=self.root,
        )


def main() -> None:
    parser = argparse.ArgumentParser(description="Генератор сценариев")
    parser.add_argument("--generate", choices=SITE_CONFIGURATIONS)
    parser.add_argument("--source-dir", type=Path)
    parser.add_argument("--output-path", type=Path)
    arguments = parser.parse_args()
    if arguments.generate is not None:
        configuration = SITE_CONFIGURATIONS[arguments.generate]
        source_directory = arguments.source_dir or configuration.source_directory
        output_path = (
            arguments.output_path or source_directory / configuration.output_filename
        )
        generator_module = load_generator_module(configuration.generator_path)
        generator_module.create_scenario(
            source_directory=source_directory,
            output_path=output_path,
        )
        return
    if arguments.source_dir is not None or arguments.output_path is not None:
        parser.error("--source-dir и --output-path требуют --generate")

    if sys.platform == "win32":
        # Отключаем растяжение готового bitmap, чтобы текст и элементы Tk были чёткими.
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            ctypes.windll.user32.SetProcessDPIAware()

    root = tk.Tk()
    try:
        generator_module = load_generator_module()
    except Exception as error:  # noqa: BLE001
        root.withdraw()
        messagebox.showerror("Ошибка запуска", str(error), parent=root)
        root.destroy()
        return

    ScenarioGeneratorApp(root, generator_module)
    root.mainloop()


if __name__ == "__main__":
    main()
