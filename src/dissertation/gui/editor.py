"""Окно редактора: структура документа, панель ввода для выбранного блока, живой предпросмотр."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from ..document_builder import BLOCK_LABELS, BlockType, InsertPosition
from .common import guarded
from .panels import BibliographyPanel, TextPanel, TitlePanel, TocPanel
from .preview import PreviewView

POSITIONS = [("auto", "Автоматически (по смыслу блока)"), ("end", "В конец документа"),
             ("after", "После выбранного в структуре"), ("section_end", "В конец выбранного раздела")]
PREVIEW_DELAY_MS = 250


class EditorView:
    def __init__(self, master, app):
        self.app = app
        self.ctl = app.ctl
        self.root = app.root
        self._iid_to_el: dict[str, object] = {}
        self._job = None

        self.frame = ttk.Frame(master, padding=8)
        paned = ttk.PanedWindow(self.frame, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # ---- слева: структура
        left = ttk.Frame(paned, padding=(0, 0, 8, 0))
        ttk.Label(left, text="Структура документа", font=("TkDefaultFont", 10, "bold")).pack(anchor="w")
        tbox = ttk.Frame(left)
        tbox.pack(fill="both", expand=True, pady=(4, 0))
        self.tree = ttk.Treeview(tbox, show="tree", selectmode="browse")
        sb = ttk.Scrollbar(tbox, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.schedule_preview())
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
        self.block_combo.bind("<<ComboboxSelected>>", lambda e: self._on_block_change())
        self.block_combo.pack(side="left", padx=(6, 16))
        ttk.Label(head, text="Куда:").pack(side="left")
        self.pos_combo = ttk.Combobox(head, state="readonly", width=32,
                                      values=[label for _, label in POSITIONS])
        self.pos_combo.current(0)
        self.pos_combo.bind("<<ComboboxSelected>>", lambda e: self.schedule_preview())
        self.pos_combo.pack(side="left", padx=6)

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
        self.root.bind("<Control-Return>", lambda e: self.on_action())

        self.text_panel = TextPanel(self)
        self.panels = {BlockType.TITLE_PAGE: TitlePanel(self), BlockType.TOC: TocPanel(self),
                       BlockType.BIBLIOGRAPHY: BibliographyPanel(self)}
        self.current = None
        self.refresh_outline()
        self._on_block_change()
        profile = self.ctl.builder.profile
        self.status(f"Профиль: {profile.name}. Ctrl+Enter — «{self.action_btn.cget('text')}».")

    # ------------------------------------------------------------ состояние
    def block(self) -> BlockType:
        return self._blocks[self.block_combo.current()]

    def status(self, msg: str) -> None:
        self.status_var.set(msg)

    def position(self) -> InsertPosition:
        mode = POSITIONS[self.pos_combo.current()][0]
        if mode in ("after", "section_end"):
            sel = self.tree.selection()
            if not sel or sel[0] not in self._iid_to_el:
                raise ValueError("Выберите элемент в структуре документа слева")
            return InsertPosition(mode, self._iid_to_el[sel[0]])
        return InsertPosition(mode)

    @guarded
    def _on_block_change(self) -> None:
        block = self.block()
        panel = self.panels.get(block, self.text_panel)
        if panel is self.text_panel:
            panel.set_block(block)
        if panel is not self.current:
            if self.current is not None:
                self.current.frame.pack_forget()
            panel.frame.pack(fill="both", expand=True)
            self.current = panel
            panel.on_shown()
        self.action_btn.configure(text=panel.action_label)
        self.pos_combo.configure(state="readonly" if panel.uses_position else "disabled")
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
    def refresh_outline(self) -> None:
        selected = self.tree.selection()
        selected_el = self._iid_to_el.get(selected[0]) if selected else None
        self.tree.delete(*self.tree.get_children())
        self._iid_to_el.clear()
        parents = {-1: ""}
        for item in self.ctl.outline():
            parent = parents.get(item.level - 1, "")
            iid = self.tree.insert(parent, "end", text=item.text or "(пусто)", open=True)
            parents[item.level] = iid
            for deeper in [k for k in parents if k > item.level]:
                del parents[deeper]
            self._iid_to_el[iid] = item.element
            if item.element is selected_el:
                self.tree.selection_set(iid)

    def after_document_change(self, msg: str | None = None) -> None:
        self.refresh_outline()
        self.app.refresh_title()
        if msg:
            self.status(msg)
        self.schedule_preview()

    # ---------------------------------------------------------------- действие
    @guarded
    def on_action(self) -> None:
        msg = self.current.perform()
        self.after_document_change(msg)
