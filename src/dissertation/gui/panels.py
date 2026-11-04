"""Панели ввода для разных типов блоков.

Каждая панель — композиция (атрибут ``frame``), а не подкласс виджета Tk:
так наши методы не могут случайно перекрыть внутренние методы Tk.

Общий интерфейс панели:
    frame              — корневой виджет;
    action_label       — подпись главной кнопки;
    uses_position      — нужен ли выбор места вставки;
    refresh_preview()  — показать предпросмотр;
    perform()          — главное действие (вставить / сохранить);
    on_shown()         — панель стала видимой (подгрузить данные из документа).
"""
from __future__ import annotations

import tkinter as tk
from dataclasses import fields as dc_fields
from tkinter import messagebox, ttk

from ..document_builder import BLOCKS_WITH_HEADING, BlockType
from ..text_cleaner import CleanOptions
from ..title_page import FORM_FIELDS, TitlePageData, preview_lines
from .common import APP_TITLE, MONO_FONT, guarded, install_plain_paste, on_text_change, scrolled_text

PARAGRAPH_MODES = [("auto", "Автоматически"), ("blank_line", "Абзацы через пустую строку"),
                   ("line", "Каждая строка — абзац")]


class _Panel:
    action_label = "Добавить в документ"
    uses_position = False

    def __init__(self, editor):
        self.editor = editor
        self.ctl = editor.ctl
        self.root = editor.root
        self.frame = ttk.Frame(editor.input_host)

    supports_edit = False

    def changed(self, *_):
        self.editor.schedule_preview()

    def on_shown(self) -> None:
        pass

    def has_input(self) -> bool:
        return False


# ============================================================ обычный текст
class TextPanel(_Panel):
    """Введение, главы, подглавы, пункты, продолжение текста, заключение, приложения."""

    uses_position = True
    supports_edit = True

    def __init__(self, editor):
        super().__init__(editor)
        self.block = BlockType.CHAPTER
        self._snapshot = ("", "")
        self._heading_allowed = True
        f = self.frame
        top = ttk.Frame(f)
        top.pack(fill="x")
        top.columnconfigure(1, weight=1)
        self.heading_label = ttk.Label(top, text="Заголовок:")
        self.heading_label.grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.heading_var = tk.StringVar()
        self.heading_var.trace_add("write", self.changed)
        self.heading_entry = ttk.Entry(top, textvariable=self.heading_var)
        self.heading_entry.grid(row=0, column=1, sticky="ew", pady=2)

        hint = ttk.Frame(f)
        hint.pack(fill="x", pady=(6, 2))
        ttk.Label(hint, text="Текст (Ctrl+V; исходное форматирование не сохраняется):",
                  foreground="#555").pack(side="left")
        ttk.Button(hint, text="Ссылка на источник…", command=self.on_cite).pack(side="right")
        ttk.Button(hint, text="Ссылка на приложение…", command=self.on_cite_appendix).pack(
            side="right", padx=4)

        box, self.text = scrolled_text(f, wrap="word", undo=True, font=MONO_FONT, height=8,
                                       relief="solid", borderwidth=1)
        box.pack(fill="both", expand=True)
        install_plain_paste(self.text, self.root)
        on_text_change(self.text, self.changed)

        opts = ttk.Frame(f)
        opts.pack(fill="x", pady=(4, 0))
        self.clean_var = tk.BooleanVar(value=True)
        self.typo_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="Очистить текст", variable=self.clean_var,
                        command=self.changed).pack(side="left")
        ttk.Checkbutton(opts, text="«Ёлочки» и тире", variable=self.typo_var,
                        command=self.changed).pack(side="left", padx=10)
        ttk.Label(opts, text="Абзацы:").pack(side="left", padx=(10, 4))
        self.mode_combo = ttk.Combobox(opts, state="readonly", width=26,
                                       values=[label for _, label in PARAGRAPH_MODES])
        self.mode_combo.current(0)
        self.mode_combo.bind("<<ComboboxSelected>>", self.changed)
        self.mode_combo.pack(side="left")

    def set_block(self, block: BlockType) -> None:
        self.block = block
        self._heading_allowed = block in BLOCKS_WITH_HEADING
        self.heading_entry.configure(state="normal" if self._heading_allowed else "disabled")
        self.heading_label.configure(
            text="Заголовок приложения:" if block == BlockType.APPENDIX else "Заголовок:")

    def _clean_options(self) -> CleanOptions:
        return CleanOptions(paragraph_mode=PARAGRAPH_MODES[self.mode_combo.current()][0],
                            typography=self.typo_var.get())

    def _args(self) -> dict:
        return dict(text=self.text.get("1.0", "end-1c"), heading=self.heading_var.get(),
                    position=self.editor.position(), clean=self.clean_var.get(),
                    options=self._clean_options())

    # ------------------------------------------------- режим редактирования
    def _values(self) -> tuple[str, str]:
        return self.heading_var.get(), self.text.get("1.0", "end-1c")

    def has_input(self) -> bool:
        h, t = self._values()
        return bool(h.strip() or t.strip())

    def has_changes(self) -> bool:
        return self._values() != self._snapshot

    def _set_text(self, value: str, editable: bool) -> None:
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", value)
        self.text.edit_modified(False)
        self.text.configure(state="normal" if editable else "disabled",
                            background="white" if editable else "#f0f0f0")

    def load_section(self, content) -> None:
        """Показать содержимое выбранного пункта для правки."""
        self.heading_entry.configure(state="normal")
        self.heading_var.set(content.heading)
        if not content.heading_editable:
            self.heading_entry.configure(state="disabled")
        self._set_text(content.body, content.body_editable)
        self._snapshot = self._values()

    def clear_input(self) -> None:
        self.heading_entry.configure(state="normal")
        self.heading_var.set("")
        self.heading_entry.configure(state="normal" if self._heading_allowed else "disabled")
        self._set_text("", True)
        self._snapshot = ("", "")

    def save_edit(self, element) -> str:
        h, t = self._values()
        res = self.ctl.update_section(element, h, t, clean=self.clean_var.get(),
                                      options=self._clean_options())
        if not res.ok:
            raise ValueError(res.message)
        self._snapshot = self._values()
        return res.message + (" · " + "; ".join(res.warnings) if res.warnings else "")

    def refresh_preview(self) -> None:
        el = self.editor.editing_el
        if el is not None:
            h, t = self._values()
            res = self.ctl.preview_edit(el, h, t, clean=self.clean_var.get(),
                                        options=self._clean_options())
        else:
            res = self.ctl.preview_block(self.block, **self._args())
        self.editor.preview.show_result(res)

    def perform(self) -> str:
        res = self.ctl.insert(self.block, **self._args())
        self.clear_input()
        msg = f"Добавлено абзацев: {len(res.paragraphs)}"
        return msg + (" · " + "; ".join(res.warnings) if res.warnings else "")

    @guarded
    def on_cite(self) -> None:
        from .dialogs import CitationDialog

        entries = [e for e in self.ctl.bib_entries() if e.key]
        if not entries:
            messagebox.showinfo(APP_TITLE, "Список литературы пуст (или его записи ещё не под "
                                "управлением программы). Добавьте источники: тип блока "
                                "«Список литературы».", parent=self.root)
            return
        dlg = CitationDialog(self.root, entries, self.ctl.profile_separator())
        if dlg.result:
            self.text.insert("insert", dlg.result)
            self.text.focus_set()


    @guarded
    def on_cite_appendix(self) -> None:
        from .dialogs import AppendixRefDialog

        entries = [e for e in self.ctl.app_entries() if e.key]
        if not entries:
            messagebox.showinfo(APP_TITLE, "Приложений пока нет. Добавьте их: тип блока «Приложение».",
                                parent=self.root)
            return
        dlg = AppendixRefDialog(self.root, entries)
        if dlg.result:
            self.text.insert("insert", dlg.result)
            self.text.focus_set()


# =========================================================== титульный лист
class TitlePanel(_Panel):
    action_label = "Сохранить титульный лист"

    def __init__(self, editor):
        super().__init__(editor)
        canvas = tk.Canvas(self.frame, highlightthickness=0)
        sb = ttk.Scrollbar(self.frame, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        inner.columnconfigure(1, weight=1)

        self.vars: dict[str, object] = {}
        for row, (key, label, multiline) in enumerate(FORM_FIELDS):
            ttk.Label(inner, text=label + ":").grid(row=row, column=0, sticky="nw", padx=(0, 8), pady=3)
            if key == "manuscript":
                var = tk.BooleanVar(value=True)
                ttk.Checkbutton(inner, variable=var, command=self.changed).grid(
                    row=row, column=1, sticky="w", pady=3)
            elif multiline:
                var = tk.Text(inner, height=2, wrap="word", relief="solid", borderwidth=1,
                              font=("TkDefaultFont", 10), undo=True)
                var.grid(row=row, column=1, sticky="ew", pady=3)
                install_plain_paste(var, self.root)
                on_text_change(var, self.changed)
            else:
                var = tk.StringVar()
                var.trace_add("write", self.changed)
                ttk.Entry(inner, textvariable=var).grid(row=row, column=1, sticky="ew", pady=3)
            self.vars[key] = var
        ttk.Label(inner, foreground="#555", wraplength=520, justify="left",
                  text="Перевод строки внутри поля сохраняется. Повторное сохранение формы "
                       "заменяет титульный лист целиком; при открытии документа форма "
                       "заполняется из него.").grid(row=len(FORM_FIELDS), column=0, columnspan=2,
                                                    sticky="w", pady=(8, 0))

    def on_shown(self) -> None:
        self.fill(self.ctl.title_data())

    def fill(self, data: TitlePageData) -> None:
        for f in dc_fields(TitlePageData):
            var = self.vars.get(f.name)
            value = getattr(data, f.name)
            if isinstance(var, tk.Text):
                var.delete("1.0", "end")
                var.insert("1.0", value)
                var.edit_modified(False)
            elif var is not None:
                var.set(value)

    def data(self) -> TitlePageData:
        values = {}
        for key, var in self.vars.items():
            values[key] = var.get("1.0", "end-1c") if isinstance(var, tk.Text) else var.get()
        return TitlePageData(**values)

    def refresh_preview(self) -> None:
        self.editor.preview.show_title(preview_lines(self.data()))

    def perform(self) -> str:
        self.ctl.set_title_page(self.data())
        return "Титульный лист сохранён в документе"


# ========================================================== оглавление
class TocPanel(_Panel):
    action_label = "Вставить / обновить оглавление"

    def __init__(self, editor):
        super().__init__(editor)
        ttk.Label(self.frame, justify="left", wraplength=600, text=(
            "Оглавление вставляется после титульного листа как поле Word и строится по "
            "заголовкам (главы, подглавы, пункты, структурные элементы).\n\n"
            "• Программа обновляет список пунктов при каждом изменении структуры.\n"
            "• Номера страниц ставит Word: при открытии он предложит обновить поля — ответьте "
            "«Да» (или F9 на оглавлении).\n"
            "• В LibreOffice: Сервис → Обновить → Все указатели."
        )).pack(anchor="w", pady=8)

    def refresh_preview(self) -> None:
        self.editor.preview.show_result(self.ctl.preview_block(BlockType.TOC))

    def perform(self) -> str:
        res = self.ctl.insert(BlockType.TOC)
        return "; ".join(res.warnings) or "Оглавление вставлено"


# ===================================================== список литературы
class BibliographyPanel(_Panel):
    action_label = "Добавить записи в список"

    def __init__(self, editor):
        super().__init__(editor)
        f = self.frame
        cols = ("num", "refs", "text")
        tbox = ttk.Frame(f)
        tbox.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tbox, columns=cols, show="headings", selectmode="browse", height=7)
        for col, title, width, stretch in (("num", "№", 40, False), ("refs", "Ссылок", 60, False),
                                           ("text", "Запись", 500, True)):
            self.tree.heading(col, text=title)
            self.tree.column(col, width=width, stretch=stretch, anchor="w" if col == "text" else "center")
        sb = ttk.Scrollbar(tbox, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        btns = ttk.Frame(f)
        btns.pack(fill="x", pady=4)
        for text, cmd in (("↑", lambda: self._move(-1)), ("↓", lambda: self._move(1)),
                          ("Удалить", self.on_remove), ("Сортировать А→Я", lambda: self._sort("alpha")),
                          ("По порядку упоминания", lambda: self._sort("citation"))):
            ttk.Button(btns, text=text, command=cmd).pack(side="left", padx=(0, 4))
        self.save_btn = ttk.Button(btns, text="Сохранить правку записи", command=self.on_update)
        self.save_btn.pack(side="right")

        ttk.Label(f, foreground="#555", text="Новые записи — по одной в строке (номера "
                  "не нужны, программа нумерует сама). Выберите запись в таблице, чтобы "
                  "исправить её текст:").pack(anchor="w", pady=(4, 2))
        box, self.text = scrolled_text(f, wrap="word", undo=True, font=MONO_FONT, height=4,
                                       relief="solid", borderwidth=1)
        box.pack(fill="x")
        install_plain_paste(self.text, self.root)
        on_text_change(self.text, self.changed)
        self._iid_key: dict[str, str] = {}

    def on_shown(self) -> None:
        legacy = sum(1 for e in self.ctl.bib_entries() if e.key is None)
        if legacy and messagebox.askyesno(
                APP_TITLE, f"В списке литературы {legacy} записей, созданных не этой программой.\n"
                           "Взять их под управление (автонумерация, сортировка, ссылки)? "
                           "Текст записей сохранится, оформление — по ГОСТ.", parent=self.root):
            self.ctl.bib_adopt()
            self.editor.after_document_change()
        self.reload()

    def reload(self, select_key: str | None = None) -> None:
        self.tree.delete(*self.tree.get_children())
        self._iid_key.clear()
        cites = self.ctl.bib_citations()
        for i, e in enumerate(self.ctl.bib_entries(), 1):
            iid = self.tree.insert("", "end", values=(i, cites.get(e.key, 0) if e.key else "—", e.text))
            if e.key:
                self._iid_key[iid] = e.key
                if e.key == select_key:
                    self.tree.selection_set(iid)
                    self.tree.see(iid)

    def _selected_key(self) -> str | None:
        sel = self.tree.selection()
        if not sel:
            return None
        key = self._iid_key.get(sel[0])
        if key is None:
            messagebox.showinfo(APP_TITLE, "Эта запись не под управлением программы — откройте "
                                "панель заново и согласитесь её принять.", parent=self.root)
        return key

    def _on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        text = self.tree.item(sel[0], "values")[2]
        self.text.delete("1.0", "end")
        self.text.insert("1.0", text)
        self.text.edit_modified(False)

    def _done(self, msg: str, key: str | None = None) -> None:
        self.reload(select_key=key)
        self.editor.after_document_change(msg)

    @guarded
    def _move(self, delta: int) -> None:
        key = self._selected_key()
        if key:
            self.ctl.bib_move(key, delta)
            self._done("Порядок изменён, номера и ссылки обновлены", key)

    @guarded
    def _sort(self, mode: str) -> None:
        self.ctl.bib_sort(mode)
        self._done("Список отсортирован " + ("по алфавиту" if mode == "alpha" else "по порядку упоминания")
                   + ", ссылки в тексте обновлены")

    @guarded
    def on_remove(self) -> None:
        key = self._selected_key()
        if not key:
            return
        refs = self.ctl.bib_citations().get(key, 0)
        warn = f"\nНа неё {refs} ссылок в тексте — они превратятся в «[?]»." if refs else ""
        if messagebox.askyesno(APP_TITLE, "Удалить запись?" + warn, parent=self.root):
            self.ctl.bib_remove(key)
            self.text.delete("1.0", "end")
            self._done("Запись удалена")

    @guarded
    def on_update(self) -> None:
        key = self._selected_key()
        text = " ".join(self.text.get("1.0", "end-1c").split())
        if key and text:
            self.ctl.bib_update(key, text)
            self._done("Запись исправлена", key)

    def refresh_preview(self) -> None:
        text = "" if self.tree.selection() else self.text.get("1.0", "end-1c")
        self.editor.preview.show_result(self.ctl.preview_block(BlockType.BIBLIOGRAPHY, text=text))

    def perform(self) -> str:
        if self.tree.selection():
            self.tree.selection_remove(*self.tree.selection())
        lines = self.text.get("1.0", "end-1c")
        res = self.ctl.insert(BlockType.BIBLIOGRAPHY, text=lines)
        self.text.delete("1.0", "end")
        self.reload()
        return f"Добавлено записей: {len(res.paragraphs)}"


# ============================================================== приложения
class AppendixPanel(_Panel):
    """Приложения списком (как литература): буквы А, Б, В… по порядку; ссылки {@P1}."""

    def __init__(self, editor):
        super().__init__(editor)
        f = self.frame
        tbox = ttk.Frame(f)
        tbox.pack(fill="x")
        self.tree = ttk.Treeview(tbox, columns=("letter", "refs", "title"), show="headings",
                                 selectmode="browse", height=5)
        for col, title, width, stretch in (("letter", "Буква", 60, False), ("refs", "Ссылок", 60, False),
                                           ("title", "Заголовок приложения", 500, True)):
            self.tree.heading(col, text=title)
            self.tree.column(col, width=width, stretch=stretch, anchor="w" if col == "title" else "center")
        sb = ttk.Scrollbar(tbox, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="x", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        btns = ttk.Frame(f)
        btns.pack(fill="x", pady=4)
        for text, cmd in (("Новое приложение", self.on_new), ("↑", lambda: self._move(-1)),
                          ("↓", lambda: self._move(1)), ("Удалить", self.on_remove),
                          ("По порядку упоминания", self.on_sort)):
            ttk.Button(btns, text=text, command=cmd).pack(side="left", padx=(0, 4))

        row = ttk.Frame(f)
        row.pack(fill="x", pady=(4, 2))
        row.columnconfigure(1, weight=1)
        ttk.Label(row, text="Заголовок:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.heading_var = tk.StringVar()
        self.heading_var.trace_add("write", self.changed)
        ttk.Entry(row, textvariable=self.heading_var).grid(row=0, column=1, sticky="ew")
        self.mode_label = ttk.Label(f, foreground="#555")
        self.mode_label.pack(anchor="w")
        box, self.text = scrolled_text(f, wrap="word", undo=True, font=MONO_FONT, height=6,
                                       relief="solid", borderwidth=1)
        box.pack(fill="both", expand=True)
        install_plain_paste(self.text, self.root)
        on_text_change(self.text, self.changed)
        self._iid_el: dict[str, object] = {}
        self._iid_key: dict[str, str | None] = {}
        self.selected_el = None

    @property
    def action_label(self) -> str:
        return "Сохранить приложение" if self.selected_el is not None else "Добавить приложение"

    def has_input(self) -> bool:
        return bool(self.heading_var.get().strip() or self.text.get("1.0", "end-1c").strip())

    def on_shown(self) -> None:
        self.reload()

    def reload(self, select_el=None) -> None:
        self.tree.delete(*self.tree.get_children())
        self._iid_el.clear()
        self._iid_key.clear()
        refs = self.ctl.app_citations()
        target = None
        for e in self.ctl.app_entries():
            iid = self.tree.insert("", "end", values=(e.letter, refs.get(e.key, 0) if e.key else "—",
                                                      e.title or "(без заголовка)"))
            self._iid_el[iid] = e.element
            self._iid_key[iid] = e.key
            if select_el is not None and e.element is select_el:
                target = iid
        if target:
            self.tree.selection_set(target)
            self.tree.see(target)
            self._load(target)
        elif select_el is None and self.selected_el is not None:
            self._set_new_mode()

    def select_element(self, el) -> None:
        self.reload(select_el=el)

    def _set_new_mode(self) -> None:
        self.selected_el = None
        self.heading_var.set("")
        self.text.configure(state="normal", background="white")
        self.text.delete("1.0", "end")
        self.mode_label.configure(text="Новое приложение — получит следующую букву")
        self.editor._update_buttons()

    def _load(self, iid) -> None:
        el = self._iid_el[iid]
        content = self.ctl.section_content(el)
        self.selected_el = el
        self.heading_var.set(content.heading)
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", content.body)
        self.text.edit_modified(False)
        if not content.body_editable:
            self.text.configure(state="disabled", background="#f0f0f0")
        letter = self.tree.item(iid, "values")[0]
        key = self._iid_key.get(iid)
        self.mode_label.configure(text=f"Приложение {letter}" + (f" · ссылка в тексте: {{@{key}}}" if key else "")
                                  + ("" if content.body_editable else " · " + content.reason))
        self.editor._update_buttons()

    def _on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if sel and self._iid_el.get(sel[0]) is not self.selected_el:
            self._load(sel[0])
            self.changed()

    @guarded
    def on_new(self) -> None:
        if self.tree.selection():
            self.tree.selection_remove(*self.tree.selection())
        self._set_new_mode()
        self.changed()

    @guarded
    def _move(self, delta: int) -> None:
        if self.selected_el is None:
            return
        res = self.ctl.move_item(self.selected_el, delta)
        el = self.selected_el
        self.reload(select_el=el)
        self.editor.after_document_change(res.message)

    @guarded
    def on_sort(self) -> None:
        n = self.ctl.app_sort_by_mention()
        self.reload(select_el=self.selected_el)
        self.editor.after_document_change("Приложения расположены в порядке ссылок в тексте" if n
                                          else "Порядок уже соответствует ссылкам в тексте")

    @guarded
    def on_remove(self) -> None:
        if self.selected_el is None:
            return
        if not messagebox.askyesno(APP_TITLE, "Удалить " + self.ctl.describe_item(self.selected_el) + "?",
                                   parent=self.root):
            return
        res = self.ctl.delete_item(self.selected_el)
        self._set_new_mode()
        self.reload()
        self.editor.after_document_change(res.message)

    def refresh_preview(self) -> None:
        h, t = self.heading_var.get(), self.text.get("1.0", "end-1c")
        if self.selected_el is not None:
            res = self.ctl.preview_edit(self.selected_el, h, t)
        else:
            res = self.ctl.preview_block(BlockType.APPENDIX, text=t, heading=h)
        self.editor.preview.show_result(res)

    def perform(self) -> str:
        h, t = self.heading_var.get(), self.text.get("1.0", "end-1c")
        if self.selected_el is not None:
            el = self.selected_el
            res = self.ctl.update_section(el, h, t)
            if not res.ok:
                raise ValueError(res.message)
            self.reload(select_el=el)
            return res.message + (" · " + "; ".join(res.warnings) if res.warnings else "")
        if not h.strip():
            raise ValueError("Введите заголовок приложения")
        before = {id(e.element) for e in self.ctl.app_entries()}
        res = self.ctl.insert(BlockType.APPENDIX, text=t, heading=h)
        new = next((e.element for e in self.ctl.app_entries() if id(e.element) not in before), None)
        self.reload(select_el=new)
        return "Приложение добавлено" + (" · " + "; ".join(res.warnings) if res.warnings else "")
