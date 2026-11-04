"""Главное окно приложения.

Экран 1 — главное меню: «Выбор документа», «Создание документа», «Загрузить образец».
Экран 2 — редактор: структура документа слева, справа — тип блока, заголовок,
место вставки и окно ввода «сырого» текста без инструментов форматирования.

Вся логика — в AppController; здесь только виджеты и обработка событий.
"""
from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ..controller import AppController
from .common import APP_TITLE, IS_WIN, guarded, install_entry_shortcuts
from .dialogs import DOCX_TYPES, OpenDialog, ProfileDialog, ReportDialog
from .editor import EditorView


def open_externally(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class App:
    def __init__(self, root: tk.Tk, controller: AppController | None = None):
        self.root = root
        self.ctl = controller or AppController()
        self.root.title(APP_TITLE)
        self.root.geometry("1100x720")
        self.root.minsize(900, 560)
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)
        style = ttk.Style(self.root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Big.TButton", font=("Segoe UI", 12), padding=(20, 12))
        self._build_menu()
        install_entry_shortcuts(self.root)
        self.container = ttk.Frame(self.root)
        self.container.pack(fill="both", expand=True)
        self.editor: EditorView | None = None
        self.show_main_menu()

    # ------------------------------------------------------------- меню окна
    def _build_menu(self) -> None:
        m = tk.Menu(self.root)
        f = tk.Menu(m, tearoff=False)
        f.add_command(label="Создание документа…", command=self.on_new, accelerator="Ctrl+N")
        f.add_command(label="Выбор документа…", command=self.on_open, accelerator="Ctrl+O")
        f.add_separator()
        f.add_command(label="Сохранить", command=self.on_save, accelerator="Ctrl+S")
        f.add_command(label="Сохранить как…", command=self.on_save_as)
        f.add_command(label="Открыть в Word", command=self.on_open_in_word)
        f.add_separator()
        f.add_command(label="Загрузить образец оформления…", command=self.on_load_template)
        f.add_command(label="Главное меню", command=self.show_main_menu)
        f.add_separator()
        f.add_command(label="Выход", command=self.on_exit)
        m.add_cascade(label="Файл", menu=f)
        h = tk.Menu(m, tearoff=False)
        h.add_command(label="Текущий профиль оформления", command=self.on_show_profile)
        m.add_cascade(label="Справка", menu=h)
        self.root.config(menu=m)
        # bind по keycode: на русской раскладке keysym другой, и <Control-s> не срабатывает.
        self.root.bind_all("<Control-KeyPress>", self._on_ctrl_key)

    def _on_ctrl_key(self, event):
        key = {83: self.on_save, 78: self.on_new, 79: self.on_open}.get(event.keycode) if IS_WIN else None
        if key is None and event.keysym.lower() in ("s", "n", "o"):
            key = {"s": self.on_save, "n": self.on_new, "o": self.on_open}[event.keysym.lower()]
        if key:
            key()
            return "break"
        return None

    def _clear(self) -> None:
        for w in self.container.winfo_children():
            w.destroy()
        self.editor = None

    # ------------------------------------------------------------ экраны
    def show_main_menu(self) -> None:
        if not self._confirm_discard():
            return
        self._clear()
        frame = ttk.Frame(self.container, padding=40)
        frame.place(relx=0.5, rely=0.45, anchor="center")
        ttk.Label(frame, text=APP_TITLE, font=("Segoe UI", 20, "bold")).pack(pady=(0, 6))
        ttk.Label(frame, text="Оформление текста по ГОСТ 7.32-2017 / ГОСТ Р 7.0.11-2011",
                  foreground="#555").pack(pady=(0, 24))
        for text, cmd in (("Выбор документа", self.on_open), ("Создание документа", self.on_new),
                          ("Загрузить образец оформления", self.on_load_template)):
            ttk.Button(frame, text=text, style="Big.TButton", command=cmd, width=32).pack(pady=6)

    def show_editor(self) -> None:
        self._clear()
        self.editor = EditorView(self.container, self)
        self.editor.frame.pack(fill="both", expand=True)
        self.refresh_title()

    def refresh_title(self) -> None:
        path = self.ctl.path
        name = path.name if path else "Новый документ"
        star = " *" if self.ctl.dirty else ""
        profile = self.ctl.builder.profile.name if self.ctl.builder else ""
        self.root.title(f"{name}{star} — {profile} — {APP_TITLE}")

    # ------------------------------------------------------------ действия
    def _confirm_discard(self) -> bool:
        if self.ctl.builder is None or not self.ctl.dirty:
            return True
        ans = messagebox.askyesnocancel(APP_TITLE, "Сохранить изменения в документе?", parent=self.root)
        if ans is None:
            return False
        if ans:
            return bool(self.on_save())
        return True

    @guarded
    def on_new(self) -> None:
        if not self._confirm_discard():
            return
        dlg = ProfileDialog(self.root, self.ctl, title="Создание документа",
                            prompt="Шаблон оформления нового документа:")
        if dlg.result is None:
            return
        self.ctl.new_document(dlg.result)
        self.show_editor()

    @guarded
    def on_open(self) -> None:
        if not self._confirm_discard():
            return
        path = filedialog.askopenfilename(parent=self.root, title="Выбор документа",
                                          filetypes=DOCX_TYPES)
        if not path:
            return
        dlg = OpenDialog(self.root, self.ctl)
        if dlg.result is None:
            return
        profile_key, enforce = dlg.result
        self.ctl.open_document(path, profile_key=profile_key, enforce_styles=enforce)
        self.show_editor()

    @guarded
    def on_load_template(self) -> None:
        path = filedialog.askopenfilename(parent=self.root, title="Образец оформления (.docx)",
                                          filetypes=DOCX_TYPES)
        if not path:
            return
        profile, saved = self.ctl.load_template(path)
        ReportDialog(self.root, f"Профиль «{profile.name}»",
                     "\n".join(profile.extraction_report) + f"\n\nСохранён: {saved}\n"
                     "Его можно выбрать при создании документа.")

    @guarded
    def on_save(self) -> bool:
        if self.ctl.builder is None:
            return False
        if self.ctl.path is None:
            return bool(self.on_save_as())
        self.ctl.save()
        self.refresh_title()
        if self.editor:
            self.editor.status(f"Сохранено: {self.ctl.path}")
        return True

    @guarded
    def on_save_as(self) -> bool:
        if self.ctl.builder is None:
            return False
        path = filedialog.asksaveasfilename(parent=self.root, title="Сохранить как",
                                            defaultextension=".docx", filetypes=DOCX_TYPES)
        if not path:
            return False
        self.ctl.save(path)
        self.refresh_title()
        if self.editor:
            self.editor.status(f"Сохранено: {self.ctl.path}")
        return True

    @guarded
    def on_open_in_word(self) -> None:
        if self.ctl.builder is None:
            return
        if self.ctl.dirty or self.ctl.path is None:
            if not self.on_save():
                return
        open_externally(self.ctl.path)

    @guarded
    def on_show_profile(self) -> None:
        if self.ctl.builder is None:
            messagebox.showinfo(APP_TITLE, "Документ не открыт", parent=self.root)
            return
        p = self.ctl.builder.profile
        b = p.body
        lines = [
            f"Профиль: {p.name}  ({p.source})",
            f"Поля, мм: левое {p.page.margin_left_mm:g}, правое {p.page.margin_right_mm:g}, "
            f"верхнее {p.page.margin_top_mm:g}, нижнее {p.page.margin_bottom_mm:g}",
            f"Основной текст: {b.font_name} {b.font_size_pt:g} пт, интервал {b.line_spacing:g}, "
            f"отступ {b.first_line_indent_cm:g} см, {b.alignment}",
            f"Нумерация заголовков: {'да — ' + p.numbering.chapter if p.numbering.enabled else 'нет'}",
            f"Номера страниц: {'да, ' + p.page_numbering.position if p.page_numbering.enabled else 'нет'}",
        ]
        if p.extraction_report:
            lines += ["", "Извлечено из образца:"] + p.extraction_report
        ReportDialog(self.root, "Профиль оформления", "\n".join(lines))

    def on_exit(self) -> None:
        if self._confirm_discard():
            self.root.destroy()


def run() -> None:
    root = tk.Tk()
    try:  # чёткие шрифты на Windows с масштабированием
        from ctypes import windll  # type: ignore[attr-defined]
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001
        pass
    App(root)
    root.mainloop()
