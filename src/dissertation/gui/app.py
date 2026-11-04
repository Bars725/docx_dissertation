"""Главное окно приложения.

Экран 1 — главное меню: «Выбор документа», «Создание документа», «Загрузить образец».
Экран 2 — редактор: структура документа слева, справа — тип блока, заголовок,
место вставки и окно ввода «сырого» текста без инструментов форматирования.

Вся логика — в AppController; здесь только виджеты и обработка событий.
"""
from __future__ import annotations

import functools
import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ..controller import AppController
from ..document_builder import BLOCK_LABELS, BLOCKS_WITH_HEADING, BlockType, InsertPosition
from ..text_cleaner import CleanOptions

APP_TITLE = "Диссертация по ГОСТ"
IS_WIN = sys.platform.startswith("win")
DOCX_TYPES = [("Документ Word", "*.docx")]
PARAGRAPH_MODES = [("auto", "Автоматически"), ("blank_line", "Абзацы через пустую строку"),
                   ("line", "Каждая строка — абзац")]
POSITIONS = [("auto", "Автоматически (по смыслу блока)"), ("end", "В конец документа"),
             ("after", "После выбранного элемента"), ("section_end", "В конец выбранного раздела")]


def guarded(method):
    """Ошибки обработчиков показываем пользователю, а не в консоль."""
    @functools.wraps(method)
    def wrapper(self, *args, **kwargs):
        try:
            return method(self, *args, **kwargs)
        except Exception as exc:  # noqa: BLE001 — интерфейс не должен падать
            messagebox.showerror(APP_TITLE, f"{type(exc).__name__}: {exc}", parent=self.root)
            return None
    return wrapper


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
        self.container = ttk.Frame(self.root)
        self.container.pack(fill="both", expand=True)
        self.editor: EditorFrame | None = None
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
        self.editor = EditorFrame(self.container, self)
        self.editor.pack(fill="both", expand=True)
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


class EditorFrame(ttk.Frame):
    def __init__(self, master, app: App):
        super().__init__(master, padding=8)
        self.app = app
        self.ctl = app.ctl
        self.root = app.root
        self._iid_to_el: dict[str, object] = {}

        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # ---- слева: структура
        left = ttk.Frame(paned, padding=(0, 0, 8, 0))
        ttk.Label(left, text="Структура документа", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.tree = ttk.Treeview(left, show="tree", selectmode="browse")
        sb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, pady=(4, 0))
        sb.pack(side="left", fill="y", pady=(4, 0))
        paned.add(left, weight=1)

        # ---- справа: ввод
        right = ttk.Frame(paned)
        paned.add(right, weight=3)
        form = ttk.Frame(right)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Тип блока:").grid(row=0, column=0, sticky="w", pady=3)
        self.block_var = tk.StringVar()
        self._blocks = list(BLOCK_LABELS.items())
        self.block_combo = ttk.Combobox(form, state="readonly", textvariable=self.block_var,
                                        values=[label for _, label in self._blocks])
        self.block_combo.grid(row=0, column=1, sticky="ew", pady=3)
        self.block_combo.current(2)  # «Глава»
        self.block_combo.bind("<<ComboboxSelected>>", lambda e: self._on_block_change())

        ttk.Label(form, text="Заголовок:").grid(row=1, column=0, sticky="w", pady=3)
        self.heading_var = tk.StringVar()
        self.heading_entry = ttk.Entry(form, textvariable=self.heading_var)
        self.heading_entry.grid(row=1, column=1, sticky="ew", pady=3)

        ttk.Label(form, text="Куда вставить:").grid(row=2, column=0, sticky="nw", pady=3)
        pos = ttk.Frame(form)
        pos.grid(row=2, column=1, sticky="w")
        self.pos_var = tk.StringVar(value="auto")
        for i, (key, label) in enumerate(POSITIONS):
            ttk.Radiobutton(pos, text=label, value=key, variable=self.pos_var).grid(
                row=i // 2, column=i % 2, sticky="w", padx=(0, 16))

        ttk.Label(right, text="Текст (вставьте Ctrl+V; любое исходное форматирование будет "
                              "заменено оформлением по ГОСТ):", foreground="#555").pack(anchor="w", pady=(8, 2))
        text_frame = ttk.Frame(right)
        text_frame.pack(fill="both", expand=True)
        self.text = tk.Text(text_frame, wrap="word", undo=True, font=("Consolas", 11),
                            relief="solid", borderwidth=1)
        tsb = ttk.Scrollbar(text_frame, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=tsb.set)
        self.text.pack(side="left", fill="both", expand=True)
        tsb.pack(side="left", fill="y")
        self._install_plain_paste()

        opts = ttk.Frame(right)
        opts.pack(fill="x", pady=(6, 0))
        self.clean_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Очистить текст (переносы, лишние пробелы)",
                        variable=self.clean_var).pack(side="left")
        self.typo_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="«Ёлочки» и тире", variable=self.typo_var).pack(side="left", padx=12)
        ttk.Label(opts, text="Абзацы:").pack(side="left", padx=(12, 4))
        self.mode_var = tk.StringVar(value=PARAGRAPH_MODES[0][1])
        ttk.Combobox(opts, state="readonly", width=28, textvariable=self.mode_var,
                     values=[label for _, label in PARAGRAPH_MODES]).pack(side="left")

        btns = ttk.Frame(right)
        btns.pack(fill="x", pady=8)
        ttk.Button(btns, text="Добавить в документ", command=self.on_insert).pack(side="right")
        ttk.Button(btns, text="Предпросмотр абзацев", command=self.on_preview).pack(side="right", padx=6)
        ttk.Button(btns, text="Очистить поле", command=lambda: self.text.delete("1.0", "end")).pack(side="left")

        self.status_var = tk.StringVar()
        ttk.Label(self, textvariable=self.status_var, foreground="#555").pack(fill="x", pady=(4, 0))

        self._on_block_change()
        self.refresh_outline()
        if self.ctl.builder and self.ctl.builder.profile.extraction_report:
            self.status(f"Профиль: {self.ctl.builder.profile.name} (Справка → профиль оформления)")

    # ----------------------------------------------------------- вставка текста
    def _install_plain_paste(self) -> None:
        """Только простой текст: Tk и так не принимает RTF/HTML, но ещё убираем \\r
        и даём рабочий Ctrl+V на русской раскладке (keycode 86)."""
        def paste(_event=None):
            try:
                data = self.root.clipboard_get()
            except tk.TclError:
                return "break"
            try:
                self.text.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            self.text.insert("insert", data.replace("\r\n", "\n").replace("\r", "\n"))
            return "break"

        def on_ctrl(event):
            if (IS_WIN and event.keycode == 86) or event.keysym.lower() == "v":
                return paste()
            if (IS_WIN and event.keycode == 65) or event.keysym.lower() == "a":
                self.text.tag_add("sel", "1.0", "end")
                return "break"
            return None

        self.text.bind("<Control-KeyPress>", on_ctrl)
        self.text.bind("<<Paste>>", paste)
        self.text.bind("<Shift-Insert>", paste)
        menu = tk.Menu(self.text, tearoff=False)
        menu.add_command(label="Вставить", command=paste)
        menu.add_command(label="Очистить", command=lambda: self.text.delete("1.0", "end"))
        self.text.bind("<Button-3>", lambda e: menu.tk_popup(e.x_root, e.y_root))

    # ------------------------------------------------------------- состояние
    def status(self, msg: str) -> None:
        self.status_var.set(msg)

    def _block(self) -> BlockType:
        return self._blocks[self.block_combo.current()][0]

    def _on_block_change(self) -> None:
        enabled = self._block() in BLOCKS_WITH_HEADING
        self.heading_entry.configure(state="normal" if enabled else "disabled")

    def _options(self) -> CleanOptions:
        mode = next(k for k, label in PARAGRAPH_MODES if label == self.mode_var.get())
        return CleanOptions(paragraph_mode=mode, typography=self.typo_var.get())

    def _position(self) -> InsertPosition:
        mode = self.pos_var.get()
        if mode in ("after", "section_end"):
            sel = self.tree.selection()
            if not sel:
                raise ValueError("Выберите элемент в структуре документа слева")
            return InsertPosition(mode, self._iid_to_el[sel[0]])
        return InsertPosition(mode)

    def refresh_outline(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self._iid_to_el.clear()
        if self.ctl.builder is None:
            return
        parents = {-1: ""}
        for i, item in enumerate(self.ctl.outline()):
            parent = parents.get(item.level - 1, "")
            iid = self.tree.insert(parent, "end", text=item.text or "(пусто)", open=True)
            parents[item.level] = iid
            for deeper in [k for k in parents if k > item.level]:
                del parents[deeper]
            self._iid_to_el[iid] = item.element

    # --------------------------------------------------------------- действия
    @guarded
    def on_preview(self) -> None:
        paras = self.ctl.preview(self._block(), self.text.get("1.0", "end-1c"),
                                 clean=self.clean_var.get(), options=self._options())
        body = "\n\n".join(f"¶{i}  {p}" for i, p in enumerate(paras, 1)) or "(пусто)"
        ReportDialog(self.root, f"Предпросмотр: {len(paras)} абз.", body)

    @guarded
    def on_insert(self) -> None:
        res = self.ctl.insert(self._block(), text=self.text.get("1.0", "end-1c"),
                              heading=self.heading_var.get(), position=self._position(),
                              clean=self.clean_var.get(), options=self._options())
        self.text.delete("1.0", "end")
        self.heading_var.set("")
        self.refresh_outline()
        self.app.refresh_title()
        msg = f"Добавлено абзацев: {len(res.paragraphs)}"
        if res.warnings:
            msg += " · " + "; ".join(res.warnings)
        self.status(msg)


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

    def __init__(self, master, ctl: AppController, title: str, prompt: str):
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

    def __init__(self, master, ctl: AppController):
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
        t = tk.Text(self.body, wrap="word", width=90, height=24, font=("Segoe UI", 10))
        t.insert("1.0", text)
        t.configure(state="disabled")
        t.pack(fill="both", expand=True)
        ttk.Button(self.body, text="OK", command=self.destroy).pack(anchor="e", pady=(8, 0))
        self.run()


def run() -> None:
    root = tk.Tk()
    try:  # чёткие шрифты на Windows с масштабированием
        from ctypes import windll  # type: ignore[attr-defined]
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001
        pass
    App(root)
    root.mainloop()
