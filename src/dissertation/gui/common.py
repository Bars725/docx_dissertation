"""Общие помощники GUI."""
from __future__ import annotations

import functools
import sys
import tkinter as tk
from tkinter import messagebox

APP_TITLE = "Диссертация по ГОСТ"
IS_WIN = sys.platform.startswith("win")
UI_FONT = ("Segoe UI", 10) if IS_WIN else ("TkDefaultFont", 10)
MONO_FONT = ("Consolas", 11) if IS_WIN else ("TkFixedFont", 11)


def guarded(method):
    """Ошибки обработчиков показываем пользователю, а не в консоль (нужен self.root)."""
    @functools.wraps(method)
    def wrapper(self, *args, **kwargs):
        try:
            return method(self, *args, **kwargs)
        except Exception as exc:  # noqa: BLE001 — интерфейс не должен падать
            messagebox.showerror(APP_TITLE, f"{type(exc).__name__}: {exc}", parent=self.root)
            return None
    return wrapper


def install_plain_paste(text: tk.Text, root) -> None:
    """Вставка только простого текста; Ctrl+V/Ctrl+A работают и на русской раскладке."""
    def paste(_event=None):
        try:
            data = root.clipboard_get()
        except tk.TclError:
            return "break"
        try:
            text.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        text.insert("insert", data.replace("\r\n", "\n").replace("\r", "\n"))
        return "break"

    def on_ctrl(event):
        if (IS_WIN and event.keycode == 86) or event.keysym.lower() == "v":
            return paste()
        if (IS_WIN and event.keycode == 65) or event.keysym.lower() == "a":
            text.tag_add("sel", "1.0", "end")
            return "break"
        if (IS_WIN and event.keycode == 67) or event.keysym.lower() == "c":
            text.event_generate("<<Copy>>")
            return "break"
        if (IS_WIN and event.keycode == 88) or event.keysym.lower() == "x":
            text.event_generate("<<Cut>>")
            return "break"
        if (IS_WIN and event.keycode == 90) or event.keysym.lower() == "z":
            try:
                text.edit_undo()
            except tk.TclError:
                pass
            return "break"
        return None

    text.bind("<Control-KeyPress>", on_ctrl)
    text.bind("<<Paste>>", paste)
    text.bind("<Shift-Insert>", paste)
    menu = tk.Menu(text, tearoff=False)
    menu.add_command(label="Вставить", command=paste)
    menu.add_command(label="Копировать", command=lambda: text.event_generate("<<Copy>>"))
    menu.add_command(label="Очистить", command=lambda: text.delete("1.0", "end"))
    text.bind("<Button-3>", lambda e: menu.tk_popup(e.x_root, e.y_root))


def on_text_change(text: tk.Text, callback) -> None:
    """Вызывать callback при любом изменении содержимого tk.Text."""
    def handler(_event=None):
        if text.edit_modified():
            text.edit_modified(False)
            callback()
    text.bind("<<Modified>>", handler)


def scrolled_text(parent, **kw):
    """(frame, text) — Text с вертикальной прокруткой."""
    from tkinter import ttk

    frame = ttk.Frame(parent)
    t = tk.Text(frame, **kw)
    sb = ttk.Scrollbar(frame, orient="vertical", command=t.yview)
    t.configure(yscrollcommand=sb.set)
    t.pack(side="left", fill="both", expand=True)
    sb.pack(side="left", fill="y")
    return frame, t
