# -*- coding: utf-8 -*-
"""Современный интерфейс приложения сводного отчёта."""

from __future__ import annotations

import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from . import theme
from .report_builder import ReportValidationError, build_summary_report, validate_export

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class FileCard(ctk.CTkFrame):
    """Карточка выбора одной выгрузки."""

    def __init__(
        self,
        master,
        kind: str,
        title: str,
        hint: str,
        on_changed,
        **kwargs,
    ):
        super().__init__(
            master,
            fg_color=theme.CARD,
            corner_radius=16,
            border_width=1,
            border_color=theme.BORDER,
            **kwargs,
        )
        self.kind = kind
        self.on_changed = on_changed
        self.path: Path | None = None

        self.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(self, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 6))
        top.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            top,
            text=title,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=15, weight="bold"),
            text_color=theme.TEXT,
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        self.badge = ctk.CTkLabel(
            top,
            text="не выбран",
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=12, weight="bold"),
            text_color=theme.MUTED,
            fg_color=theme.ACCENT_SOFT,
            corner_radius=999,
            padx=10,
            pady=4,
        )
        self.badge.grid(row=0, column=1, sticky="e")

        ctk.CTkLabel(
            self,
            text=hint,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=12),
            text_color=theme.MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 8))

        row = ctk.CTkFrame(self, fg_color="transparent")
        row.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 16))
        row.grid_columnconfigure(0, weight=1)

        self.path_label = ctk.CTkLabel(
            row,
            text="Файл не выбран",
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=12),
            text_color=theme.MUTED,
            anchor="w",
            wraplength=620,
            justify="left",
        )
        self.path_label.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        ctk.CTkButton(
            row,
            text="Обзор…",
            width=110,
            height=34,
            corner_radius=10,
            fg_color=theme.NAVY,
            hover_color=theme.NAVY_HOVER,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=13, weight="bold"),
            command=self.browse,
        ).grid(row=0, column=1, sticky="e")

    def browse(self) -> None:
        path = filedialog.askopenfilename(
            title=theme.FILE_LABELS[self.kind],
            filetypes=[
                ("Excel", "*.xlsx *.xlsm *.xls"),
                ("Все файлы", "*.*"),
            ],
        )
        if path:
            self.set_path(Path(path))

    def set_path(self, path: Path) -> None:
        self.path = path
        self.path_label.configure(text=str(path), text_color=theme.TEXT)
        try:
            validate_export(path, self.kind)
        except ReportValidationError as exc:
            self.badge.configure(
                text="ошибка",
                text_color=theme.DANGER,
                fg_color=theme.DANGER_BG,
            )
            self.path_label.configure(text=f"{path.name} — {exc}", text_color=theme.DANGER)
            self.on_changed()
            return

        self.badge.configure(
            text="OK",
            text_color=theme.SUCCESS,
            fg_color=theme.SUCCESS_BG,
        )
        self.on_changed()

    def is_ready(self) -> bool:
        if self.path is None:
            return False
        try:
            validate_export(self.path, self.kind)
            return True
        except ReportValidationError:
            return False


class MetricChip(ctk.CTkFrame):
    def __init__(self, master, title: str, **kwargs):
        super().__init__(
            master,
            fg_color=theme.ACCENT_SOFT,
            corner_radius=14,
            border_width=0,
            **kwargs,
        )
        self.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            self,
            text=title,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=11),
            text_color=theme.MUTED,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 0))
        self.value = ctk.CTkLabel(
            self,
            text="—",
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=22, weight="bold"),
            text_color=theme.NAVY,
            anchor="w",
        )
        self.value.grid(row=1, column=0, sticky="ew", padx=14, pady=(2, 14))

    def set_value(self, text: str, color: str = theme.NAVY) -> None:
        self.value.configure(text=text, text_color=color)


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Мотивирующий мониторинг — сводный отчёт")
        self.geometry(theme.WINDOW_SIZE)
        self.minsize(*theme.WINDOW_MIN)
        self.configure(fg_color=theme.BG)

        self.output_dir = tk.StringVar(value="")
        self.output_name = tk.StringVar(value=theme.DEFAULT_OUTPUT_NAME)
        self._last_output: Path | None = None
        self._busy = False

        self._build_layout()
        self._try_autoload()
        self._refresh_generate_state()

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        root = ctk.CTkScrollableFrame(self, fg_color=theme.BG, corner_radius=0)
        root.grid(row=0, column=0, sticky="nsew")
        root.grid_columnconfigure(0, weight=1)

        # Header
        header = ctk.CTkFrame(
            root,
            fg_color=theme.NAVY,
            corner_radius=20,
        )
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(24, 16))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="Мотивирующий мониторинг",
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=26, weight="bold"),
            text_color="#FFFFFF",
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=24, pady=(22, 4))

        ctk.CTkLabel(
            header,
            text="Загрузите три выгрузки Excel — программа сформирует сводный отчёт "
            "по региону и районам: своевременность, соответствие ТМ, СанПиН.",
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=13),
            text_color="#D7E6F8",
            anchor="w",
            wraplength=900,
            justify="left",
        ).grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 22))

        # File cards
        cards = ctk.CTkFrame(root, fg_color="transparent")
        cards.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 12))
        cards.grid_columnconfigure(0, weight=1)

        self.card_tim = FileCard(
            cards,
            kind="timeliness",
            title="1. Своевременность размещения меню",
            hint="Ожидаемые столбцы: Регион, Район, Пищеблок, Процент Своевременности",
            on_changed=self._on_files_changed,
        )
        self.card_tim.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        self.card_tm = FileCard(
            cards,
            kind="typical_menu",
            title="2. Соответствие типовому меню",
            hint="Ожидаемые столбцы: Пищеблок Наименование, Район, Дата, Число Несоблюдений, "
            "Процент Соблюдения, Замечания",
            on_changed=self._on_files_changed,
        )
        self.card_tm.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        self.card_err = FileCard(
            cards,
            kind="sanpin",
            title="3. Ошибки по СанПиН",
            hint="Ожидаемые столбцы: Район, Пищеблок, Дата, Прием Пищи, Ошибка",
            on_changed=self._on_files_changed,
        )
        self.card_err.grid(row=2, column=0, sticky="ew", pady=(0, 10))

        # Output settings
        out_card = ctk.CTkFrame(
            root,
            fg_color=theme.CARD,
            corner_radius=16,
            border_width=1,
            border_color=theme.BORDER,
        )
        out_card.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 12))
        out_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            out_card,
            text="Куда сохранить отчёт",
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=15, weight="bold"),
            text_color=theme.TEXT,
            anchor="w",
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=18, pady=(16, 10))

        ctk.CTkLabel(
            out_card,
            text="Папка",
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=12),
            text_color=theme.MUTED,
        ).grid(row=1, column=0, sticky="w", padx=(18, 8), pady=4)

        self.dir_entry = ctk.CTkEntry(
            out_card,
            textvariable=self.output_dir,
            height=36,
            corner_radius=10,
            border_color=theme.BORDER,
            fg_color="#F8FAFC",
            text_color=theme.TEXT,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=12),
        )
        self.dir_entry.grid(row=1, column=1, sticky="ew", pady=4)

        ctk.CTkButton(
            out_card,
            text="Выбрать…",
            width=110,
            height=36,
            corner_radius=10,
            fg_color=theme.TEAL,
            hover_color=theme.TEAL_HOVER,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=13, weight="bold"),
            command=self._pick_output_dir,
        ).grid(row=1, column=2, padx=(10, 18), pady=4)

        ctk.CTkLabel(
            out_card,
            text="Имя файла",
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=12),
            text_color=theme.MUTED,
        ).grid(row=2, column=0, sticky="w", padx=(18, 8), pady=(4, 16))

        self.name_entry = ctk.CTkEntry(
            out_card,
            textvariable=self.output_name,
            height=36,
            corner_radius=10,
            border_color=theme.BORDER,
            fg_color="#F8FAFC",
            text_color=theme.TEXT,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=12),
        )
        self.name_entry.grid(row=2, column=1, columnspan=2, sticky="ew", padx=(0, 18), pady=(4, 16))

        # Actions
        actions = ctk.CTkFrame(root, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", padx=24, pady=(4, 8))
        actions.grid_columnconfigure(0, weight=1)

        self.generate_btn = ctk.CTkButton(
            actions,
            text="Сформировать сводный отчёт",
            height=48,
            corner_radius=14,
            fg_color=theme.NAVY,
            hover_color=theme.NAVY_HOVER,
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=16, weight="bold"),
            command=self._generate,
        )
        self.generate_btn.grid(row=0, column=0, sticky="ew")

        self.progress = ctk.CTkProgressBar(
            actions,
            height=8,
            corner_radius=999,
            progress_color=theme.TEAL,
            fg_color="#E2E8F0",
        )
        self.progress.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        self.progress.set(0)
        self.progress.grid_remove()

        self.status = ctk.CTkLabel(
            actions,
            text="Выберите три выгрузки, чтобы начать",
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=12),
            text_color=theme.MUTED,
            anchor="w",
        )
        self.status.grid(row=2, column=0, sticky="ew", pady=(10, 0))

        # Results
        self.result_card = ctk.CTkFrame(
            root,
            fg_color=theme.CARD,
            corner_radius=16,
            border_width=1,
            border_color=theme.BORDER,
        )
        self.result_card.grid(row=4, column=0, sticky="ew", padx=24, pady=(8, 28))
        self.result_card.grid_columnconfigure((0, 1, 2, 3), weight=1)

        ctk.CTkLabel(
            self.result_card,
            text="Результат по региону",
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=15, weight="bold"),
            text_color=theme.TEXT,
            anchor="w",
        ).grid(row=0, column=0, columnspan=4, sticky="w", padx=18, pady=(16, 10))

        self.chip_composite = MetricChip(self.result_card, "Комплексный показатель")
        self.chip_tim = MetricChip(self.result_card, "Своевременность")
        self.chip_tm = MetricChip(self.result_card, "Соответствие ТМ")
        self.chip_err = MetricChip(self.result_card, "Ошибки СанПиН")

        self.chip_composite.grid(row=1, column=0, sticky="ew", padx=(18, 6), pady=(0, 8))
        self.chip_tim.grid(row=1, column=1, sticky="ew", padx=6, pady=(0, 8))
        self.chip_tm.grid(row=1, column=2, sticky="ew", padx=6, pady=(0, 8))
        self.chip_err.grid(row=1, column=3, sticky="ew", padx=(6, 18), pady=(0, 8))

        self.result_note = ctk.CTkLabel(
            self.result_card,
            text="После генерации здесь появятся ключевые цифры и путь к файлу.",
            font=ctk.CTkFont(family=theme.FONT_FAMILY, size=12),
            text_color=theme.MUTED,
            anchor="w",
            justify="left",
            wraplength=900,
        )
        self.result_note.grid(row=2, column=0, columnspan=4, sticky="ew", padx=18, pady=(4, 8))

        btns = ctk.CTkFrame(self.result_card, fg_color="transparent")
        btns.grid(row=3, column=0, columnspan=4, sticky="w", padx=18, pady=(0, 16))

        self.open_file_btn = ctk.CTkButton(
            btns,
            text="Открыть Excel",
            width=140,
            height=36,
            corner_radius=10,
            fg_color=theme.TEAL,
            hover_color=theme.TEAL_HOVER,
            state="disabled",
            command=self._open_file,
        )
        self.open_file_btn.pack(side="left", padx=(0, 8))

        self.open_folder_btn = ctk.CTkButton(
            btns,
            text="Открыть папку",
            width=140,
            height=36,
            corner_radius=10,
            fg_color=theme.NAVY,
            hover_color=theme.NAVY_HOVER,
            state="disabled",
            command=self._open_folder,
        )
        self.open_folder_btn.pack(side="left")

    def _try_autoload(self) -> None:
        """Подставляет файлы из соседней папки, если они уже лежат рядом с проектом."""
        parent = Path(__file__).resolve().parents[2]
        candidates = {
            "timeliness": [
                parent / "Своевременность.xlsx",
                Path(r"C:\Users\Kheda\Downloads\Своевременность.xlsx"),
            ],
            "typical_menu": [parent / "Соотв_типовому меню.xlsx"],
            "sanpin": [parent / "Кол-во ошибок.xlsx"],
        }
        mapping = {
            "timeliness": self.card_tim,
            "typical_menu": self.card_tm,
            "sanpin": self.card_err,
        }
        found_any = False
        for kind, paths in candidates.items():
            for p in paths:
                if p.exists():
                    mapping[kind].set_path(p)
                    found_any = True
                    break
        if found_any and not self.output_dir.get():
            self.output_dir.set(str(parent))

    def _on_files_changed(self) -> None:
        if not self.output_dir.get():
            for card in (self.card_tim, self.card_tm, self.card_err):
                if card.path is not None:
                    self.output_dir.set(str(card.path.parent))
                    break
        self._refresh_generate_state()

    def _refresh_generate_state(self) -> None:
        ready = all(c.is_ready() for c in (self.card_tim, self.card_tm, self.card_err))
        if self._busy:
            self.generate_btn.configure(state="disabled")
            return
        self.generate_btn.configure(state="normal" if ready else "disabled")
        if ready:
            self.status.configure(
                text="Все файлы проверены — можно формировать отчёт",
                text_color=theme.SUCCESS,
            )
        else:
            missing = []
            for card, label in (
                (self.card_tim, "своевременность"),
                (self.card_tm, "типовое меню"),
                (self.card_err, "СанПиН"),
            ):
                if not card.is_ready():
                    missing.append(label)
            self.status.configure(
                text="Ожидаются корректные файлы: " + ", ".join(missing),
                text_color=theme.MUTED,
            )

    def _pick_output_dir(self) -> None:
        path = filedialog.askdirectory(title="Папка для сводного отчёта")
        if path:
            self.output_dir.set(path)

    def _resolve_output_path(self) -> Path:
        folder = self.output_dir.get().strip()
        name = self.output_name.get().strip() or theme.DEFAULT_OUTPUT_NAME
        if not folder:
            for card in (self.card_tm, self.card_tim, self.card_err):
                if card.path is not None:
                    folder = str(card.path.parent)
                    break
        if not folder:
            folder = str(Path.home() / "Downloads")
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
        return Path(folder) / name

    def _generate(self) -> None:
        if self._busy:
            return
        if not all(c.is_ready() for c in (self.card_tim, self.card_tm, self.card_err)):
            messagebox.showwarning("Не хватает файлов", "Загрузите и проверьте все три выгрузки.")
            return

        out_path = self._resolve_output_path()
        self._busy = True
        self.generate_btn.configure(state="disabled", text="Формирование…")
        self.progress.grid()
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        self.status.configure(text="Считаем показатели и собираем Excel…", text_color=theme.NAVY)
        self.update_idletasks()

        args = (
            self.card_tim.path,
            self.card_tm.path,
            self.card_err.path,
            out_path,
        )

        def worker() -> None:
            try:
                result = build_summary_report(*args)
                self.after(0, lambda: self._on_success(result))
            except ReportValidationError as exc:
                self.after(0, lambda: self._on_error(str(exc)))
            except Exception as exc:  # noqa: BLE001
                self.after(0, lambda: self._on_error(f"Непредвиденная ошибка: {exc}"))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, result: dict) -> None:
        self._busy = False
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(1)
        self.generate_btn.configure(text="Сформировать сводный отчёт")
        self._refresh_generate_state()

        self._last_output = Path(result["output_path"])
        self.chip_composite.set_value(f"{result['composite']:.2f}%", theme.NAVY)
        self.chip_tim.set_value(f"{result['timeliness_mean']:.2f}%", theme.SUCCESS)
        self.chip_tm.set_value(f"{result['typical_menu_mean']:.2f}%", theme.TEAL)
        self.chip_err.set_value(str(result["sanpin_errors"]), theme.WARNING)

        top = ", ".join(result["top_districts"])
        attention = ", ".join(result["attention_districts"])
        self.result_note.configure(
            text=(
                f"Регион: {result['region']}  •  районов: {result['districts']}\n"
                f"Лучшие: {top}. Требуют внимания: {attention}.\n"
                f"Файл: {result['output_path']}"
            ),
            text_color=theme.TEXT,
        )
        self.status.configure(text="Отчёт успешно сформирован", text_color=theme.SUCCESS)
        self.open_file_btn.configure(state="normal")
        self.open_folder_btn.configure(state="normal")
        messagebox.showinfo("Готово", f"Сводный отчёт сохранён:\n{result['output_path']}")

    def _on_error(self, message: str) -> None:
        self._busy = False
        self.progress.stop()
        self.progress.grid_remove()
        self.generate_btn.configure(text="Сформировать сводный отчёт")
        self._refresh_generate_state()
        self.status.configure(text=message, text_color=theme.DANGER)
        messagebox.showerror("Ошибка", message)

    def _open_file(self) -> None:
        if self._last_output and self._last_output.exists():
            os.startfile(self._last_output)  # noqa: S606

    def _open_folder(self) -> None:
        if self._last_output and self._last_output.exists():
            os.startfile(self._last_output.parent)  # noqa: S606


def run_app() -> None:
    app = App()
    app.mainloop()
