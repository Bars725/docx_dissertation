"""Окно редактора: структура документа, панель ввода для выбранного блока, живой предпросмотр.

Выбор пункта в структуре загружает его содержимое для правки («Сохранить изменения»).
«Новый блок» очищает поля — дальше «Добавить в документ» вставит новый фрагмент
в место, указанное в «Куда» (можно относительно выбранного пункта).
Под структурой: ▲/▼ (Ctrl+↑/↓) — перемещение в пределах своего уровня, «Удалить» (Delete).
"""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from ..document_builder import BLOCK_LABELS, BlockType, InsertPosition
from ..outline import Kind
from .common import APP_TITLE, guarded
from .panels import BibliographyPanel, TextPanel, TitlePanel, TocPanel
from .preview import PreviewView

POSITIONS = [("auto", "Автоматически (по смыслу блока)"), ("end", "В конец документа"),
             ("after", "После выбранного в структуре"), ("section_end", "В конец выбранного раздела")]
PREVIEW_DELAY_MS = 250
KIND_TO_BLOCK = {Kind.TITLE: BlockType.TITLE_PAGE, Kind.TOC: BlockType.TOC,
                 Kind.INTRODUCTION: BlockType.INTRODUCTION, Kind.CHAPTER: BlockType.CHAPTER,
                 Kind.SUBCHAPTER: BlockType.SUBCHAPTER, Kind.SECTION: BlockType.SECTION,
                 Kind.CONCLUSION: BlockType.CONCLUSION, Kind.BIBLIOGRAPHY: BlockType.BIBLIOGRAPHY,
                 Kind.APPENDIX: BlockType.APPENDIX, Kind.OTHER_STRUCTURAL: BlockType.BODY}


class EditorView:
    def __init__(self, master, app):
        self.app = app
        self.ctl = app.ctl
        self.root = app.root
        self._iid_to_el: dict[str, object] = {}
        self._iid_kind: dict[str, Kind] = {}
        self._job = None
        self._suppress_select = False
        self._selected_iid = None
        self.editing_el = None

        self.frame = ttk.Frame(master, padding=8)
        paned = ttk.PanedWindow(self.frame, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # ---- слева: структура + кнопки
        left = ttk.Frame(paned, padding=(0, 0, 8, 0))
        ttk.Label(left, text="Структура документа", font=("TkDefaultFont", 10, "bold")).pack(anchor="w")
        tools = ttk.Frame(left)
        tools.pack(side="bottom", fill="x", pady=(6, 0))
        ttk.Button(tools, text="▲ Выше", width=9, command=lambda: self.on_move(-1)).pack(side="left")
        ttk.Button(tools, text="▼ Ниже", width=9, command=lambda: self.on_move(1)).pack(side="left", padx=4)
        ttk.Button(tools, text="✕ Удалить", width=11, command=self.on_delete).pack(side="right")
        ttk.Label(left, text="Ctrl+↑/↓ — переместить, Delete — удалить", foreground="#777").pack(
            side="bottom", anchor="w")
        tbox = ttk.Frame(left)
        tbox.pack(fill="both", expand=True, pady=(4, 0))
        self.tree = ttk.Treeview(tbox, show="tree", selectmode="browse")
        sb = ttk.Scrollbar(tbox, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Delete>", lambda e: (self.on_delete(), "break")[1])
        self.tree.bind("<Control-Up>", lambda e: (self.on_move(-1), "break")[1])
        self.tree.bind("<Control-Down>", lambda e: (self.on_move(1), "break")[1])
        paned.add(left, weight=1)

        # ---- справа
        right = ttk.Frame(paned)
        paned.add(right, weight=3)
        head = ttk.Frame(right)
        head.pack(fill="x")
        ttk.Label(head, text="Тип блока:").pack(side="left")
        self._blocks = list(BLOCK_LABELS)
        self.block_combo = ttk.Combobox(head, state="readonly", width=30,
                                        values=[BLOCK_LABELS[b] for b in self._blocks])
        self.block_combo.current(self._blocks.index(BlockType.CHAPTER))
        self.block_combo.bind("<<ComboboxSelected>>", lambda e: self._on_block_change(user=True))
        self.block_combo.pack(side="left", padx=(6, 16))
        ttk.Label(head, text="Куда:").pack(side="left")
        self.pos_combo = ttk.Combobox(head, state="readonly", width=32,
                                      values=[label for _, label in POSITIONS])
        self.pos_combo.current(0)
        self.pos_combo.bind("<<ComboboxSelected>>", lambda e: self.schedule_preview())
        self.pos_combo.pack(side="left", padx=6)

        modebar = ttk.Frame(right)
        modebar.pack(fill="x", pady=(6, 0))
        self.mode_var = tk.StringVar(value="Новый блок")
        ttk.Label(modebar, textvariable=self.mode_var, foreground="#1a5fb4").pack(side="left")
        self.new_btn = ttk.Button(modebar, text="Новый блок (очистить поля)", command=self.on_new_block)
        self.new_btn.pack(side="right")

        vpaned = ttk.PanedWindow(right, orient="vertical")
        vpaned.pack(fill="both", expand=True, pady=(6, 0))
        self.input_host = ttk.Frame(vpaned)
        vpaned.add(self.input_host, weight=3)
        self.preview = PreviewView(vpaned, self.ctl.builder.profile)
        vpaned.add(self.preview.frame, weight=2)

        bottom = ttk.Frame(right)
        bottom.pack(fill="x", pady=(6, 0))
        self.status_var = tk.StringVar()
        ttk.Label(bottom, textvariable=self.status_var, foreground="#555").pack(side="left", fill="x",
                                                                                expand=True)
        self.action_btn = ttk.Button(bottom, text="Добавить в документ", command=self.on_action)
        self.action_btn.pack(side="right")
        self.save_btn = ttk.Button(bottom, text="Сохранить изменения", command=self.on_save_edit)
        self.save_btn.pack(side="right", padx=6)
        self.root.bind("<Control-Return>", lambda e: self.on_primary())

        self.text_panel = TextPanel(self)
        self.panels = {BlockType.TITLE_PAGE: TitlePanel(self), BlockType.TOC: TocPanel(self),
                       BlockType.BIBLIOGRAPHY: BibliographyPanel(self)}
        self.current = None
        self.refresh_outline()
        self._on_block_change()
        self.status(f"Профиль: {self.ctl.builder.profile.name}. Выберите пункт слева, чтобы "
                    "увидеть и изменить его текст.")

    # ------------------------------------------------------------ состояние
    def block(self) -> BlockType:
        return self._blocks[self.block_combo.current()]

    def status(self, msg: str) -> None:
        self.status_var.set(msg)

    def selected_element(self):
        sel = self.tree.selection()
        return self._iid_to_el.get(sel[0]) if sel else None

    def position(self) -> InsertPosition:
        mode = POSITIONS[self.pos_combo.current()][0]
        if mode in ("after", "section_end"):
            el = self.selected_element()
            if el is None:
                raise ValueError("Выберите элемент в структуре документа слева")
            return InsertPosition(mode, el)
        return InsertPosition(mode)

    def _set_editing(self, element, label: str = "") -> None:
        self.editing_el = element
        if element is None:
            self.mode_var.set("Новый блок — «Добавить в документ» вставит его в место из «Куда»")
        else:
            self.mode_var.set(f"Редактирование: «{label}» — «Сохранить изменения» заменит текст пункта")
        self._update_buttons()

    def _update_buttons(self) -> None:
        panel = self.current
        editing = self.editing_el is not None and panel is self.text_panel
        self.save_btn.configure(state="normal" if editing else "disabled")
        self.new_btn.configure(state="normal" if panel is self.text_panel else "disabled")
        if panel is not None:
            self.action_btn.configure(text=panel.action_label)
            self.pos_combo.configure(state="readonly" if panel.uses_position else "disabled")

    def _show_block(self, block: BlockType) -> None:
        self.block_combo.current(self._blocks.index(block))
        panel = self.panels.get(block, self.text_panel)
        if panel is self.text_panel:
            panel.set_block(block)
        if panel is not self.current:
            if self.current is not None:
                self.current.frame.pack_forget()
            panel.frame.pack(fill="both", expand=True)
            self.current = panel
            panel.on_shown()
        self._update_buttons()

    @guarded
    def _on_block_change(self, user: bool = False) -> None:
        if user and self.editing_el is not None:
            self._set_editing(None)       # смена типа вручную — значит, новый блок
        self._show_block(self.block())
        if self.editing_el is None:
            self._set_editing(None)
        self.refresh_preview()

    # --------------------------------------------------------- выбор в дереве
    @guarded
    def _on_tree_select(self, _event=None) -> None:
        if self._suppress_select:
            return
        sel = self.tree.selection()
        if not sel or sel[0] == self._selected_iid:
            return
        iid = sel[0]
        panel = self.text_panel
        dirty_edit = self.editing_el is not None and panel.has_changes()
        dirty_new = self.editing_el is None and self.current is panel and panel.has_input()
        if dirty_edit and not messagebox.askyesno(
                APP_TITLE, "Изменения в текущем пункте не сохранены. Отменить их?", parent=self.root):
            self._reselect(self._selected_iid)
            return
        if dirty_new and not messagebox.askyesno(
                APP_TITLE, "Заменить введённый текст содержимым выбранного пункта?\n"
                           "(«Нет» — оставить текст и использовать пункт как место вставки)",
                parent=self.root):
            self._selected_iid = iid
            self.schedule_preview()
            return
        self._selected_iid = iid
        self.load_item(iid)

    def _reselect(self, iid) -> None:
        self._suppress_select = True
        try:
            if iid and iid in self._iid_to_el:
                self.tree.selection_set(iid)
        finally:
            self._suppress_select = False

    def load_item(self, iid) -> None:
        el, kind = self._iid_to_el[iid], self._iid_kind[iid]
        block = KIND_TO_BLOCK.get(kind, BlockType.BODY)
        self._show_block(block)
        if self.current is self.text_panel:
            content = self.ctl.section_content(el)
            self.text_panel.load_section(content)
            self._set_editing(el, self.tree.item(iid, "text"))
            if not content.body_editable:
                self.status("Текст пункта только для чтения: " + content.reason)
        else:
            self._set_editing(None)
            self.mode_var.set(f"«{self.tree.item(iid, 'text')}» — правка в панели ниже")
        self.refresh_preview()

    # ------------------------------------------------------------ предпросмотр
    def schedule_preview(self) -> None:
        if self._job is not None:
            self.root.after_cancel(self._job)
        self._job = self.root.after(PREVIEW_DELAY_MS, self.refresh_preview)

    def refresh_preview(self) -> None:
        self._job = None
        if self.current is None:
            return
        try:
            self.current.refresh_preview()
        except Exception as exc:  # noqa: BLE001 — предпросмотр не должен мешать вводу
            self.preview.show_message(str(exc), header="Предпросмотр недоступен")

    # ------------------------------------------------------------- структура
    def refresh_outline(self, select_el=None) -> None:
        target = select_el if select_el is not None else self.selected_element()
        self._suppress_select = True
        try:
            self.tree.delete(*self.tree.get_children())
            self._iid_to_el.clear()
            self._iid_kind.clear()
            self._selected_iid = None
            parents = {-1: ""}
            for item in self.ctl.outline():
                parent = parents.get(item.level - 1, "")
                iid = self.tree.insert(parent, "end", text=item.text or "(пусто)", open=True)
                parents[item.level] = iid
                for deeper in [k for k in parents if k > item.level]:
                    del parents[deeper]
                self._iid_to_el[iid] = item.element
                self._iid_kind[iid] = item.kind
                if target is not None and item.element is target:
                    self.tree.selection_set(iid)
                    self.tree.see(iid)
                    self._selected_iid = iid
        finally:
            self._suppress_select = False
        if self.editing_el is not None and self.editing_el not in self._iid_to_el.values():
            self._set_editing(None)

    def after_document_change(self, msg: str | None = None, select_el=None) -> None:
        self.refresh_outline(select_el)
        self.app.refresh_title()
        if msg:
            self.status(msg)
        self.schedule_preview()

    # ---------------------------------------------------------------- действия
    def on_primary(self) -> None:
        """Ctrl+Enter: сохранить правку, если редактируем пункт, иначе — добавить."""
        if self.editing_el is not None and self.current is self.text_panel:
            self.on_save_edit()
        else:
            self.on_action()

    @guarded
    def on_action(self) -> None:
        msg = self.current.perform()
        if self.current is self.text_panel:
            self._set_editing(None)
        self.after_document_change(msg)

    @guarded
    def on_save_edit(self) -> None:
        if self.editing_el is None:
            return
        el = self.editing_el
        msg = self.text_panel.save_edit(el)
        self.after_document_change(msg, select_el=el)
        if self._selected_iid:
            self.load_item(self._selected_iid)     # показать сохранённое (с новыми номерами)
        self.status(msg)

    @guarded
    def on_new_block(self) -> None:
        if self.editing_el is not None and self.text_panel.has_changes() and not messagebox.askyesno(
                APP_TITLE, "Изменения в текущем пункте не сохранены. Отменить их?", parent=self.root):
            return
        self.text_panel.clear_input()
        self._set_editing(None)
        self.refresh_preview()

    @guarded
    def on_move(self, delta: int) -> None:
        el = self.selected_element()
        if el is None:
            self.status("Выберите пункт в структуре")
            return
        if self.editing_el is el and self.text_panel.has_changes():
            raise ValueError("Сначала сохраните изменения в пункте (или нажмите «Новый блок»)")
        res = self.ctl.move_item(el, delta)
        if res.ok:
            self.after_document_change(res.message, select_el=el)
            if self._selected_iid and self.editing_el is el:
                self.load_item(self._selected_iid)
        else:
            self.status(res.message)

    @guarded
    def on_delete(self) -> None:
        el = self.selected_element()
        if el is None:
            self.status("Выберите пункт в структуре")
            return
        if not messagebox.askyesno(APP_TITLE, "Удалить " + self.ctl.describe_item(el) +
                                   " вместе с содержимым?", parent=self.root):
            return
        res = self.ctl.delete_item(el)
        if self.editing_el is el:
            self.text_panel.clear_input()
            self._set_editing(None)
        self.after_document_change(res.message + (" · " + "; ".join(res.warnings) if res.warnings else ""))
