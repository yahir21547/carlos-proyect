"""Common style utilities for the Tkinter GUIs."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


# Colores usados de manera consistente en toda la interfaz
ABB_COLORS = {
    "bg": "#f0f0f0",      # Fondo principal
    "fg": "#000000",      # Texto estándar
    "accent": "#d81e05",  # Botones y elementos destacados
    "highlight": "#ffff99",  # Campo resaltado durante la edición
}


def aplicar_colorimetria(widget: tk.Misc) -> None:
    """Apply the ABB color scheme to *widget* and its ttk styles."""
    style = ttk.Style(widget)
    style.theme_use("clam")
    widget.configure(bg=ABB_COLORS["bg"])

    style.configure(".", background=ABB_COLORS["bg"], foreground=ABB_COLORS["fg"])
    style.configure("TButton", background=ABB_COLORS["accent"], foreground=ABB_COLORS["fg"])
    style.configure("TEntry", fieldbackground="white")
    style.configure("Highlight.TEntry", fieldbackground=ABB_COLORS["highlight"])
    style.configure("Completed.TEntry", fieldbackground="#ccffcc")
    style.configure("ReadonlyDark.TEntry", fieldbackground="#a9a9a9")

