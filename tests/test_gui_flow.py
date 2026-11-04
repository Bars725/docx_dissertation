"""Сквозной сценарий GUI: на фейковом Tk всегда, на настоящем — если он есть (Windows/macOS).

Сценарий: новый документ → глава (живой предпросмотр с номером) → список литературы
→ ссылка в тексте → сортировка (номер ссылки меняется) → титульный лист-форма
→ оглавление → вставка главы в конец выбранного раздела → сохранение.
"""
import importlib
import sys

import pytest
from docx import Document

import fake_tk


def _purge_gui_modules():
    for name in [m for m in sys.modules if m.startswith("dissertation.gui")]:
        del sys.modules[name]


@pytest.fixture(params=["fake", "real"])
def tk_env(request, monkeypatch):
    errors = []
    _purge_gui_modules()
    if request.param == "fake":
        tk, ttk, fd, mb = fake_tk.build_modules()
        for name, mod in (("tkinter", tk), ("tkinter.ttk", ttk), ("tkinter.filedialog", fd),
                          ("tkinter.messagebox", mb)):
            monkeypatch.setitem(sys.modules, name, mod)
        mb.showerror = lambda title, message, **k: errors.append(message)
        root = tk.Tk()
    else:
        tk = pytest.importorskip("tkinter")
        try:
            root = tk.Tk()
        except tk.TclError:
            pytest.skip("нет дисплея для Tk")
        root.withdraw()
        from tkinter import messagebox

        monkeypatch.setattr(messagebox, "showerror", lambda title, message, **k: errors.append(message))
        monkeypatch.setattr(messagebox, "showinfo", lambda *a, **k: None)
        monkeypatch.setattr(messagebox, "askyesno", lambda *a, **k: True)
        monkeypatch.setattr(messagebox, "askyesnocancel", lambda *a, **k: False)
    yield root, errors, request.param
    try:
        root.destroy()
    except Exception:  # noqa: BLE001
        pass
    _purge_gui_modules()


def _find_iid(tree, text, parent=""):
    for iid in tree.get_children(parent):
        if tree.item(iid, "text") == text:
            return iid
        found = _find_iid(tree, text, iid)
        if found:
            return found
    return None


def test_gui_scenario(tk_env, tmp_path):
    root, errors, mode = tk_env
    app_mod = importlib.import_module("dissertation.gui.app")
    from dissertation.controller import AppController
    from dissertation.document_builder import BlockType as B

    ctl = AppController(profiles_dir=tmp_path / "profiles")
    app = app_mod.App(root, ctl)
    ctl.new_document(None)
    app.show_editor()
    ed = app.editor

    def select(block):
        ed.block_combo.current(ed._blocks.index(block))
        ed._on_block_change()

    def preview_text():
        if mode == "real":
            root.update()
        ed.refresh_preview()
        return ed.preview.text.get("1.0", "end")

    # --- глава с живым предпросмотром
    select(B.CHAPTER)
    tp = ed.text_panel
    tp.heading_var.set("Обзор литературы")
    tp.text.insert("end", "Текст главы, скопиро-\nванный из PDF.")
    pv = preview_text()
    assert "1 Обзор литературы" in pv
    assert "скопированный" in pv
    assert "будет вставлено" in ed.preview.header.get()
    ed.on_action()
    assert "Добавлено абзацев: 2" in ed.status_var.get(), errors

    # --- список литературы: ввод построчно, автонумерация
    select(B.BIBLIOGRAPHY)
    bp = ed.panels[B.BIBLIOGRAPHY]
    bp.text.insert("end", "Сидоров С. С. Книга. – 2020.\nАндреев А. А. Статья. – 2021.")
    assert "2. Андреев" in preview_text()
    ed.on_action()
    assert len(bp.tree.get_children()) == 2, errors

    # --- ссылка в тексте и пересчёт после сортировки
    select(B.BODY)
    tp.text.insert("end", "Подробно описано в [@R2, с. 5].")
    assert "[2, с. 5]" in preview_text()
    ed.on_action()
    select(B.BIBLIOGRAPHY)
    bp._sort("alpha")
    assert bp.tree.item(bp.tree.get_children()[0], "values")[2].startswith("Андреев")
    texts = [p.text for p in ctl.builder.document.paragraphs]
    assert "Подробно описано в [1, с. 5]." in texts

    # --- титульный лист: форма
    select(B.TITLE_PAGE)
    form = ed.panels[B.TITLE_PAGE]
    form.vars["author"].set("Иванов Иван Иванович")
    form.vars["title"].delete("1.0", "end")
    form.vars["title"].insert("1.0", "Методы оформления")
    assert "МЕТОДЫ ОФОРМЛЕНИЯ" in preview_text()
    ed.on_action()
    assert ctl.title_data().author == "Иванов Иван Иванович"

    # --- оглавление
    select(B.TOC)
    assert "1 Обзор литературы" in preview_text()
    ed.on_action()
    assert _find_iid(ed.tree, "СОДЕРЖАНИЕ")

    # --- в конец выбранного раздела
    select(B.SUBCHAPTER)
    iid = _find_iid(ed.tree, "1 Обзор литературы")
    ed.tree.selection_set(iid)
    ed.pos_combo.current(3)
    tp.heading_var.set("Методы")
    tp.text.insert("end", "Текст подглавы.")
    assert "1.1 Методы" in preview_text()
    ed.on_action()

    out = ctl.save(tmp_path / "gui.docx")
    texts = [p.text for p in Document(str(out)).paragraphs]
    assert texts.index("1.1 Методы") < texts.index("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    assert not errors, errors
