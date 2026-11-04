"""Диалоговые окна."""
from __future__ import annotations

import os
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


# Формы ссылки (ГОСТ 7.32-2017 / ГОСТ Р 2.105: слово пишется полностью — «в соответствии
# с рисунком 2»; квадратные скобки зарезервированы за ссылками на литературу).
FIGURE_FORMS = [("на рисунке N", "рисунке {ref}"), ("с рисунком N (в соответствии с…)", "рисунком {ref}"),
                ("см. рисунок N", "рисунок {ref}"), ("рисунка N", "рисунка {ref}"),
                ("к рисунку N", "рисунку {ref}"), ("(рисунок N)", "(рисунок {ref})"),
                ("(рис. N) — сокращение, если разрешает вуз", "(рис. {ref})")]
APPENDIX_FORMS = [("в приложении А", "приложении {ref}"), ("см. приложение А", "приложение {ref}"),
                  ("приложения А", "приложения {ref}"), ("(приложение А)", "(приложение {ref})")]


class RefDialog(_Dialog):
    """Выбор цели ссылки и формы слова; результат — текст с токеном, например «рисунке {@F2}».

    Слово вставляется обычным текстом, номер/буква — полем, которое обновляется само.
    """

    def __init__(self, master, title: str, hint: str, items: list[tuple[str, str]],
                 forms: list[tuple[str, str]] | None = None, form_index: int = 0):
        super().__init__(master, title)
        self.items = items
        self.forms = forms or []
        self.form_index = form_index
        ttk.Label(self.body, text=hint, foreground="#555", justify="left").pack(anchor="w")
        self.listbox = tk.Listbox(self.body, height=10, width=80, activestyle="dotbox")
        self.listbox.pack(fill="both", expand=True, pady=6)
        for label, _ in items:
            self.listbox.insert("end", label)
        if items:
            self.listbox.selection_set(0)
        if self.forms:
            box = ttk.LabelFrame(self.body, text="Форма ссылки (слово склоняется по смыслу фразы)",
                                 padding=6)
            box.pack(fill="x", pady=(0, 6))
            self.form_var = tk.IntVar(value=min(form_index, len(self.forms) - 1))
            for i, (label, _) in enumerate(self.forms):
                ttk.Radiobutton(box, text=label, value=i, variable=self.form_var).grid(
                    row=i // 2, column=i % 2, sticky="w", padx=(0, 18))
        btns = ttk.Frame(self.body)
        btns.pack(fill="x")
        ttk.Button(btns, text="Отмена", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Вставить", command=self._ok).pack(side="right", padx=6)
        self.listbox.bind("<Double-1>", lambda e: self._ok())
        self.run()

    def _ok(self) -> None:
        sel = self.listbox.curselection()
        if not sel:
            return
        token = self.items[sel[0]][1]
        if self.forms:
            self.form_index = self.form_var.get()
            token = self.forms[self.form_index][1].format(ref=token)
        self.result = token
        self.destroy()


class ImageDialog(_Dialog):
    """«Добавить изображение»: папка → список изображений с подписями и предпросмотром.

    Подпись закрепляется за файлом (хранится в настройках) и используется в
    документе как «Рисунок N – <подпись>». Результат — маркер рисунка для окна ввода.
    """

    PREVIEW = 360

    def __init__(self, master, settings):
        super().__init__(master, "Добавить изображение")
        self.resizable(True, True)
        self.settings = settings
        self.files: list[str] = []
        self._iid_path: dict[str, str] = {}
        self._photo = None
        self._current = None

        top = ttk.Frame(self.body)
        top.pack(fill="x")
        ttk.Label(top, text="Папка с изображениями:").pack(side="left")
        self.folder_var = tk.StringVar(value=settings.image_folder)
        folder_entry = ttk.Entry(top, textvariable=self.folder_var, width=60)
        folder_entry.pack(side="left", fill="x", expand=True, padx=6)
        folder_entry.bind("<Return>", lambda e: self.reload())
        ttk.Button(top, text="Обзор…", command=self.on_browse).pack(side="left")
        ttk.Button(top, text="Обновить", command=self.reload).pack(side="left", padx=(6, 0))

        mid = ttk.Frame(self.body)
        mid.pack(fill="both", expand=True, pady=8)
        left = ttk.Frame(mid)
        left.pack(side="left", fill="both", expand=True)
        self.tree = ttk.Treeview(left, columns=("file", "caption"), show="headings",
                                 selectmode="browse", height=14)
        self.tree.heading("file", text="Файл")
        self.tree.heading("caption", text="Подпись в документе")
        self.tree.column("file", width=200, stretch=False)
        self.tree.column("caption", width=320, stretch=True)
        sb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", lambda e: self._ok())

        right = ttk.Frame(mid, padding=(10, 0, 0, 0))
        right.pack(side="left", fill="y")
        self.preview_label = ttk.Label(right, text="Выберите изображение", anchor="center",
                                       relief="solid", padding=4)
        self.preview_label.pack(fill="both", expand=True)
        self.meta_var = tk.StringVar()
        ttk.Label(right, textvariable=self.meta_var, foreground="#555").pack(anchor="w", pady=(4, 0))

        cap = ttk.Frame(self.body)
        cap.pack(fill="x")
        cap.columnconfigure(1, weight=1)
        ttk.Label(cap, text="Название рисунка:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.caption_var = tk.StringVar()
        self.caption_var.trace_add("write", lambda *_: self._update_example())
        ttk.Entry(cap, textvariable=self.caption_var).grid(row=0, column=1, sticky="ew")
        ttk.Button(cap, text="Закрепить подпись", command=self.on_pin).grid(row=0, column=2, padx=(6, 0))
        ttk.Label(cap, text="Ширина, см:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        wrow = ttk.Frame(cap)
        wrow.grid(row=1, column=1, sticky="w", pady=(6, 0))
        self.width_var = tk.StringVar()
        ttk.Entry(wrow, textvariable=self.width_var, width=7).pack(side="left")
        ttk.Label(wrow, text="пусто — автоматически (не шире полей и не выше 2/3 страницы)",
                  foreground="#777").pack(side="left", padx=6)
        self.example_var = tk.StringVar()
        ttk.Label(cap, textvariable=self.example_var, foreground="#1a5fb4").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))

        btns = ttk.Frame(self.body)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Отмена", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Вставить", command=self._ok).pack(side="right", padx=6)
        self.geometry("980x640")
        self.reload()
        self.run()

    # ------------------------------------------------------------------ папка
    def on_browse(self) -> None:
        path = filedialog.askdirectory(parent=self, title="Папка с изображениями",
                                       initialdir=self.folder_var.get() or None)
        if path:
            self.folder_var.set(path)
            self.reload()

    def reload(self) -> None:
        from ..app_settings import list_images

        folder = self.folder_var.get().strip()
        self.tree.delete(*self.tree.get_children())
        self._iid_path.clear()
        self.files = list_images(folder) if folder else []
        if folder:
            self.settings.image_folder = folder
        for path in self.files:
            iid = self.tree.insert("", "end", values=(os.path.basename(path), self._caption(path)))
            self._iid_path[iid] = path
        self.meta_var.set(f"Изображений: {len(self.files)}" if folder else "Укажите папку")
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0])
            self._on_select()

    def _caption(self, path: str) -> str:
        from ..app_settings import default_title

        return self.settings.caption(path) or default_title(path)

    # ------------------------------------------------------------- выбор
    def _on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        path = self._iid_path.get(sel[0])
        if path is None or path == self._current:
            return
        self._current = path
        self.caption_var.set(self._caption(path))
        self._show_preview(path)

    def _show_preview(self, path: str) -> None:
        self._photo = None
        meta = os.path.basename(path)
        try:
            try:
                from PIL import Image, ImageTk

                with Image.open(path) as im:
                    meta += f" · {im.width}×{im.height} px"
                    im = im.copy()
                    im.thumbnail((self.PREVIEW, self.PREVIEW))
                    self._photo = ImageTk.PhotoImage(im, master=self)
            except ImportError:
                if path.lower().endswith((".png", ".gif")):
                    photo = tk.PhotoImage(master=self, file=path)
                    factor = max(1, -(-max(photo.width(), photo.height()) // self.PREVIEW))
                    self._photo = photo.subsample(factor)
                else:
                    meta += " · предпросмотр JPG/BMP/TIFF — установите Pillow"
        except Exception as exc:  # noqa: BLE001
            meta += f" · не удалось открыть: {exc}"
        if self._photo is not None:
            self.preview_label.configure(image=self._photo, text="")
        else:
            self.preview_label.configure(image="", text="Нет предпросмотра")
        self.meta_var.set(meta)

    def _update_example(self) -> None:
        from ..figures import clean_title

        title = clean_title(self.caption_var.get()) or "…"
        self.example_var.set(f"В документе: по центру рисунок, под ним — «Рисунок N – {title}»")

    # -------------------------------------------------------------- действия
    def on_pin(self) -> None:
        if not self._current:
            return
        self.settings.set_caption(self._current, self.caption_var.get())
        for iid, path in self._iid_path.items():
            if path == self._current:
                self.tree.item(iid, values=(os.path.basename(path), self._caption(path)))

    def _ok(self) -> None:
        from ..figures import FigureSpec, clean_title

        if not self._current:
            return
        title = clean_title(self.caption_var.get()) or self._caption(self._current)
        width = None
        raw = self.width_var.get().strip().replace(",", ".")
        if raw:
            try:
                width = float(raw)
            except ValueError:
                messagebox.showerror(APP_TITLE, "Ширина — число сантиметров, например 12", parent=self)
                return
        self.settings.set_caption(self._current, title)       # подпись закрепляется за файлом
        self.result = FigureSpec(self._current, title, width).token()
        self.destroy()
