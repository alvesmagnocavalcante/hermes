from __future__ import annotations

import importlib
import inspect
import pkgutil
import threading
from collections.abc import Callable
from typing import Any

import customtkinter as ctk
from tkinter import messagebox

import automations
from automations.base import Automation


def load_automations() -> list[type[Automation]]:
    found: list[type[Automation]] = []
    for module_info in pkgutil.iter_modules(automations.__path__):
        if module_info.name.startswith("_") or module_info.name == "base":
            continue
        module = importlib.import_module(f"{automations.__name__}.{module_info.name}")
        automation_class = getattr(module, "AUTOMATION_CLASS", None)
        if (
            inspect.isclass(automation_class)
            and issubclass(automation_class, Automation)
            and automation_class is not Automation
        ):
            found.append(automation_class)
    return sorted(found, key=lambda item: item.name.lower())


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Painel de Automação de Planilhas")
        self.geometry("1100x700")
        self.minsize(900, 600)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._active_automation: Automation | None = None
        self._navigation_buttons: dict[type[Automation], ctk.CTkButton] = {}
        self._sidebar_expanded = False

        self._build_sidebar()
        self._build_content()
        self._build_footer()
        self._register_automations()

    def _build_sidebar(self) -> None:
        self.sidebar = ctk.CTkFrame(self, width=145, corner_radius=0)
        self.sidebar.grid(row=0, column=0, rowspan=2, sticky="nsew")
        self.sidebar.grid_propagate(False)

        header = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(24, 4))
        self.brand_label = ctk.CTkLabel(
            header,
            text="HERMES",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        self.brand_label.pack(side="left")
        self.sidebar_toggle = ctk.CTkButton(
            header, text="☰", width=30, height=30, fg_color="transparent",
            hover_color=("gray75", "gray25"), command=self._toggle_sidebar,
        )
        self.sidebar_toggle.pack(side="right")

        self.sidebar_subtitle = ctk.CTkLabel(
            self.sidebar,
            text="Automações",
            text_color="gray70",
        )

        self.navigation = ctk.CTkFrame(self.sidebar, fg_color="transparent")

    def _toggle_sidebar(self) -> None:
        self._sidebar_expanded = not self._sidebar_expanded
        if self._sidebar_expanded:
            self.sidebar.configure(width=230)
            self.brand_label.configure(font=ctk.CTkFont(size=24, weight="bold"))
            self.sidebar_toggle.configure(text="‹")
            self.sidebar_subtitle.pack(fill="x", padx=24, pady=(0, 18), anchor="w")
            self.navigation.pack(fill="x", padx=12)
        else:
            self.navigation.pack_forget()
            self.sidebar_subtitle.pack_forget()
            self.sidebar.configure(width=145)
            self.brand_label.configure(font=ctk.CTkFont(size=20, weight="bold"))
            self.sidebar_toggle.configure(text="☰")

    def _build_content(self) -> None:
        self.content = ctk.CTkFrame(self, corner_radius=12)
        self.content.grid(row=0, column=1, padx=18, pady=(18, 9), sticky="nsew")

    def _build_footer(self) -> None:
        self.footer = ctk.CTkFrame(self, height=95, corner_radius=12)
        self.footer.grid(row=1, column=1, padx=18, pady=(9, 18), sticky="ew")
        self.footer.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(self.footer, text="Pronto", anchor="w")
        self.status_label.grid(row=0, column=0, padx=18, pady=(12, 4), sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(self.footer)
        self.progress_bar.grid(row=1, column=0, padx=18, pady=(4, 14), sticky="ew")
        self.progress_bar.set(0)

    def _register_automations(self) -> None:
        automation_classes = load_automations()
        for automation_class in automation_classes:
            button = ctk.CTkButton(
                self.navigation,
                text=automation_class.name,
                anchor="w",
                fg_color="transparent",
                command=lambda cls=automation_class: self.show_automation(cls),
            )
            button.pack(fill="x", pady=4)
            self._navigation_buttons[automation_class] = button

        if automation_classes:
            self.show_automation(automation_classes[0])
        else:
            ctk.CTkLabel(self.content, text="Nenhuma automação encontrada.").pack(pady=40)

    def show_automation(self, automation_class: type[Automation]) -> None:
        for widget in self.content.winfo_children():
            widget.destroy()

        for registered_class, button in self._navigation_buttons.items():
            button.configure(
                fg_color=("gray75", "gray25")
                if registered_class is automation_class
                else "transparent"
            )

        self._active_automation = automation_class(self, self.content)
        self._active_automation.render()
        self.set_status(f"{automation_class.name} carregada", 0)

    def set_status(self, message: str, progress: float | None = None) -> None:
        self.status_label.configure(text=message)
        if progress is not None:
            self.progress_bar.set(max(0, min(progress, 1)))

    def run_background(
        self,
        task: Callable[[], Any],
        on_success: Callable[[Any], None] | None = None,
        on_error: Callable[[], None] | None = None,
    ) -> None:
        def runner() -> None:
            try:
                result = task()
            except Exception as error:
                self.after(
                    0,
                    lambda captured_error=error: self._handle_task_error(captured_error, on_error),
                )
            else:
                if on_success:
                    self.after(0, lambda: on_success(result))

        threading.Thread(target=runner, daemon=True).start()

    def _handle_task_error(
        self,
        error: Exception,
        on_error: Callable[[], None] | None,
    ) -> None:
        if on_error:
            on_error()
        self._show_error(error)

    def report_progress(self, message: str, progress: float) -> None:
        self.after(0, lambda: self.set_status(message, progress))

    def _show_error(self, error: Exception) -> None:
        self.set_status("Falha no processamento", 0)
        messagebox.showerror("Erro", f"Não foi possível concluir a operação.\n\n{error}")


def main() -> None:
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    App().mainloop()


if __name__ == "__main__":
    main()
