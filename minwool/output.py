"""Абстракция вывода результатов расчета."""

from dataclasses import dataclass
from typing import Any, Protocol

import pandas as pd
from tkinter import messagebox


@dataclass
class CalculationOutput:
    results: pd.DataFrame
    report: str


class OutputAdapter(Protocol):
    def render(self, gui: Any, payload: CalculationOutput) -> None:
        ...

    def set_status(self, gui: Any, status: str) -> None:
        ...

    def info(self, title: str, message: str) -> None:
        ...

    def error(self, title: str, message: str) -> None:
        ...


class TkOutputAdapter:
    """Стандартный адаптер вывода в виджеты Tkinter."""

    def render(self, gui: Any, payload: CalculationOutput) -> None:
        for item in gui.tree.get_children():
            gui.tree.delete(item)
        for _, row in payload.results.iterrows():
            gui.tree.insert("", "end", values=list(row))

        gui.debug_text.delete(1.0, "end")
        gui.debug_text.insert("end", payload.report)

    def set_status(self, gui: Any, status: str) -> None:
        gui.status_var.set(status)

    def info(self, title: str, message: str) -> None:
        messagebox.showinfo(title, message)

    def error(self, title: str, message: str) -> None:
        messagebox.showerror(title, message)
