from __future__ import annotations

"""Compatibilidade opcional com a interface CustomTkinter descontinuada."""

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    import customtkinter as ctk
except ImportError:
    tk = None
    filedialog = None
    messagebox = None
    ttk = None
    ctk = None

