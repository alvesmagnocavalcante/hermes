from __future__ import annotations

from dataclasses import dataclass
from automations.legacy_ui import ctk, ttk


@dataclass(frozen=True)
class TableColumn:
    key: str
    title: str
    width: int
    anchor: str = "w"


def create_result_table(
    parent: ctk.CTkFrame,
    columns: tuple[TableColumn, ...],
    *,
    row: int,
    padx: int = 30,
    pady: tuple[int, int] = (0, 0),
) -> ttk.Treeview:
    frame = ctk.CTkFrame(parent, fg_color="transparent")
    frame.grid(row=row, column=0, padx=padx, pady=pady, sticky="nsew")
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_columnconfigure(0, weight=1)

    style = ttk.Style()
    style.theme_use("clam")
    style.configure(
        "Hermes.Treeview", background="#1a1a1a", fieldbackground="#1a1a1a",
        foreground="#e5e7eb", borderwidth=0, rowheight=25, font=("Segoe UI", 9),
    )
    style.configure(
        "Hermes.Treeview.Heading", background="#2b2b2b", foreground="#f3f4f6",
        relief="flat", font=("Segoe UI", 9, "bold"),
    )
    style.map("Hermes.Treeview", background=[("selected", "#1f6aa5")], foreground=[("selected", "white")])

    tree = ttk.Treeview(frame, columns=[column.key for column in columns], show="headings", style="Hermes.Treeview")
    for column in columns:
        tree.heading(column.key, text=column.title, anchor=column.anchor)
        tree.column(column.key, width=column.width, minwidth=60, stretch=False, anchor=column.anchor)
    vertical = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    horizontal = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vertical.grid(row=0, column=1, sticky="ns")
    horizontal.grid(row=1, column=0, sticky="ew")
    tree.tag_configure("ok", foreground="#21a67a")
    tree.tag_configure("missing", foreground="#e0a83e")
    tree.tag_configure("error", foreground="#dc5a5a")
    return tree


def clear_table(table: ttk.Treeview) -> None:
    table.delete(*table.get_children())


def result_tag(status: str) -> str:
    text = status.casefold()
    if any(value in text for value in ("conciliado", "conciliada", "no prazo", "em dia")):
        return "ok"
    if any(value in text for value in ("ausente", "não encontrada", "nao encontrada", "alerta")):
        return "missing"
    return "error"
