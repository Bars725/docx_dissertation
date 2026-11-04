"""Диалоговые окна."""
from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .common import APP_TITLE

DOCX_TYPES = [("Документ Word", "*.docx")]


class _Dialog(tk.Toplevel):
    def __init__(self, master, title: str):
        super().__init__(master)
        self.title(title)
        self.transient(master)
        self.resizable(False, False)
        self.result = None
        self.body = ttk.Frame(self, padding=16)
        self.body.pack(fill="both", expand=True)

    def run(self):
        self.grab_set()
        self.bind("<Escape>", lambda e: self.destroy())
        self.wait_window(self)


class ProfileDialog(_Dialog):
    """Выбор шаблона: встроенный ГОСТ или сохранённый профиль-образец."""

    def __init__(self, master, ctl, title: str, prompt: str):
        super().__init__(master, title)
        self.ctl = ctl
        ttk.Label(self.body, text=prompt).pack(anchor="w")
        self.listbox = tk.Listbox(self.body, height=8, width=60, activestyle="dotbox")
        self.listbox.pack(fill="both", pady=8)
        self._fill()
        row = ttk.Frame(self.body)
        row.pack(fill="x")
        ttk.Button(row, text="Загрузить образец .docx…", command=self._load).pack(side="left")
        ttk.Button(row, text="Отмена", command=self.destroy).pack(side="right")
        ttk.Button(row, text="Создать", command=self._ok).pack(side="right", padx=6)
        self.listbox.bind("<Double-1>", lambda e: self._ok())
        self.run()

    def _fill(self, select: str | None = None) -> None:
        self.items = self.ctl.available_profiles()
        self.listbox.delete(0, "end")
        for key, label in self.items:
            kind = "ГОСТ" if key.startswith("preset:") else "образец"
            self.listbox.insert("end", f"[{kind}] {label}")
        idx = next((i for i, (k, _) in enumerate(self.items) if k == select), 0)
        self.listbox.selection_set(idx)

    def _load(self) -> None:
        path = filedialog.askopenfilename(parent=self, filetypes=DOCX_TYPES)
        if not path:
            return
        try:
            _, saved = self.ctl.load_template(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, str(exc), parent=self)
            return
        self._fill(select=f"file:{saved}")

    def _ok(self) -> None:
        sel = self.listbox.curselection()
        if sel:
            self.result = self.items[sel[0]][0]
            self.destroy()


class OpenDialog(_Dialog):
    """Как оформлять новые фрагменты в открытом документе."""

    def __init__(self, master, ctl):
        super().__init__(master, "Выбор документа")
        self.ctl = ctl
        self.mode = tk.StringVar(value="own")
        ttk.Radiobutton(self.body, text="Оформлять как в самом документе (рекомендуется)",
                        value="own", variable=self.mode).pack(anchor="w")
        ttk.Radiobutton(self.body, text="Использовать профиль:", value="profile",
                        variable=self.mode).pack(anchor="w", pady=(6, 0))
        self.items = ctl.available_profiles()
        self.combo = ttk.Combobox(self.body, state="readonly", width=50,
                                  values=[label for _, label in self.items])
        self.combo.current(0)
        self.combo.pack(anchor="w", padx=24, pady=4)
        self.enforce = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.body, text="Привести стили и поля документа к профилю\n"
                                        "(изменит вид уже набранного текста)",
                        variable=self.enforce).pack(anchor="w", padx=24)
        row = ttk.Frame(self.body)
        row.pack(fill="x", pady=(12, 0))
        ttk.Button(row, text="Отмена", command=self.destroy).pack(side="right")
        ttk.Button(row, text="Открыть", command=self._ok).pack(side="right", padx=6)
        self.run()

    def _ok(self) -> None:
        if self.mode.get() == "own":
            self.result = (None, False)
        else:
            self.result = (self.items[self.combo.current()][0], self.enforce.get())
        self.destroy()


class ReportDialog(_Dialog):
    def __init__(self, master, title: str, text: str):
        super().__init__(master, title)
        self.resizable(True, True)
        t = tk.Text(self.body, wrap="word", width=90, height=24, font=("TkDefaultFont", 10))
        t.insert("1.0", text)
        t.configure(state="disabled")
        t.pack(fill="both", expand=True)
        ttk.Button(self.body, text="OK", command=self.destroy).pack(anchor="e", pady=(8, 0))
        self.run()


class CitationDialog(_Dialog):
    """Выбор источника(ов) для ссылки; результат — токен «[@R3, с. 15]» для окна ввода."""

    def __init__(self, master, entries, separator: str = "; "):
        super().__init__(master, "Ссылка на источник")
        self.resizable(True, True)
        self.entries = entries
        ttk.Label(self.body, text="Источник (Ctrl/Shift — несколько):").pack(anchor="w")
        box = ttk.Frame(self.body)
        box.pack(fill="both", expand=True, pady=4)
        self.listbox = tk.Listbox(box, height=12, width=90, selectmode="extended", activestyle="dotbox")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        for i, e in enumerate(entries, 1):
            self.listbox.insert("end", f"{i}. {e.text}")
        if entries:
            self.listbox.selection_set(0)
        row = ttk.Frame(self.body)
        row.pack(fill="x", pady=(6, 0))
        ttk.Label(row, text="Страницы (необязательно):").pack(side="left")
        self.locator = tk.StringVar()
        ttk.Entry(row, textvariable=self.locator, width=16).pack(side="left", padx=6)
        ttk.Label(row, text="например: с. 15  или  с. 15–17", foreground="#777").pack(side="left")
        btns = ttk.Frame(self.body)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Отмена", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Вставить ссылку", command=self._ok).pack(side="right", padx=6)
        self.listbox.bind("<Double-1>", lambda e: self._ok())
        self.bind("<Return>", lambda e: self._ok())
        self.run()

    def _ok(self) -> None:
        sel = self.listbox.curselection()
        if not sel:
            return
        loc = self.locator.get().strip()
        parts = [f"@{self.entries[i].key}" for i in sel]
        if loc:
            parts[-1] += f", {loc}"
        self.result = "[" + "; ".join(parts) + "]"
        self.destroy()
