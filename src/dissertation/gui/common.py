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
        redo = (IS_WIN and event.keycode == 89) or event.keysym.lower() == "y" or (
            ((IS_WIN and event.keycode == 90) or event.keysym.lower() == "z") and event.state & 0x0001)
        if redo:
            try:
                text.edit_redo()
            except tk.TclError:
                pass
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


# Коды клавиш Windows (VK): не зависят от раскладки, в отличие от keysym.
_VK = {86: "paste", 67: "copy", 88: "cut", 65: "select_all", 90: "undo", 89: "redo"}
_LATIN = {"v": "paste", "c": "copy", "x": "cut", "a": "select_all", "z": "undo", "y": "redo"}
# На латинице paste/copy/cut делает сам Tk; отмены у Entry в Tk 8.6 нет вовсе —
# undo/redo/select_all обрабатываем на любой раскладке.
_ALWAYS = {"undo", "redo", "select_all"}
_HIST: dict[str, dict] = {}


def _hist(w) -> dict:
    h = _HIST.get(str(w))
    if h is None:
        h = _HIST[str(w)] = {"undo": [], "redo": [], "last": w.get()}
    return h


def _entry_record(event) -> None:
    """После каждого изменения поля — снимок в историю (для Ctrl+Z)."""
    w = event.widget
    try:
        value = w.get()
    except tk.TclError:
        return
    h = _hist(w)
    if value != h["last"]:
        h["undo"].append(h["last"])
        del h["undo"][:-200]
        h["redo"].clear()
        h["last"] = value


def _entry_focus(event) -> None:
    """Значение поменяла программа (загрузка пункта) — начинаем историю заново."""
    w = event.widget
    try:
        value = w.get()
    except tk.TclError:
        return
    h = _hist(w)
    if value != h["last"]:
        h.update(undo=[], redo=[], last=value)


def _entry_set(w, value: str) -> None:
    w.delete(0, "end")
    w.insert(0, value)
    _hist(w)["last"] = value


def _entry_action(w, action: str, shift: bool = False) -> str:
    if action == "undo" and shift:
        action = "redo"
    if action == "paste":
        try:
            data = w.clipboard_get()
        except tk.TclError:
            return "break"
        try:
            w.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        w.insert("insert", " ".join(data.replace("\r", "").split("\n")))
        _entry_record(type("E", (), {"widget": w})())
    elif action == "copy":
        w.event_generate("<<Copy>>")
    elif action == "cut":
        w.event_generate("<<Cut>>")
    elif action == "select_all":
        w.select_range(0, "end")
        w.icursor("end")
    elif action == "undo":
        h = _hist(w)
        _entry_record(type("E", (), {"widget": w})())
        if h["undo"]:
            h["redo"].append(w.get())
            _entry_set(w, h["undo"].pop())
    elif action == "redo":
        h = _hist(w)
        if h["redo"]:
            h["undo"].append(w.get())
            _entry_set(w, h["redo"].pop())
    return "break"


def _entry_ctrl(event):
    """Ctrl+V/C/X/A/Z/Y в однострочных полях на любой раскладке."""
    keysym = event.keysym.lower()
    shift = bool(getattr(event, "state", 0) & 0x0001)
    if keysym in _LATIN:
        action = _LATIN[keysym]
        if action not in _ALWAYS:
            return None                       # штатная привязка Tk
        return _entry_action(event.widget, action, shift)
    action = _VK.get(event.keycode) if IS_WIN else None
    if action is None:
        return None
    return _entry_action(event.widget, action, shift)


def install_entry_shortcuts(root) -> None:
    """Для всех ttk.Entry/tk.Entry приложения (включая диалоги)."""
    for cls in ("TEntry", "Entry"):
        root.bind_class(cls, "<Control-KeyPress>", _entry_ctrl, add="+")
        root.bind_class(cls, "<KeyRelease>", _entry_record, add="+")
        root.bind_class(cls, "<FocusIn>", _entry_focus, add="+")


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
