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


def test_entry_paste_on_russian_layout(monkeypatch):
    """Ctrl+V в однострочном поле при русской раскладке (keysym кириллический, VK=86)."""
    tk, ttk, fd, mb = fake_tk.build_modules()
    for name, mod in (("tkinter", tk), ("tkinter.ttk", ttk), ("tkinter.filedialog", fd),
                      ("tkinter.messagebox", mb)):
        monkeypatch.setitem(sys.modules, name, mod)
    _purge_gui_modules()
    common = importlib.import_module("dissertation.gui.common")
    monkeypatch.setattr(common, "IS_WIN", True)
    entry = ttk.Entry(None, clipboard="Заголовок\r\nиз буфера")
    event = type("E", (), {"keysym": "Cyrillic_em", "keycode": 86, "widget": entry})()
    assert common._entry_ctrl(event) == "break"
    assert entry.value == "Заголовок из буфера"
    latin = type("E", (), {"keysym": "v", "keycode": 86, "widget": entry})()
    assert common._entry_ctrl(latin) is None      # на латинице — штатная вставка Tk
    _purge_gui_modules()


def test_entry_undo_redo(monkeypatch):
    tk, ttk, fd, mb = fake_tk.build_modules()
    for name, mod in (("tkinter", tk), ("tkinter.ttk", ttk), ("tkinter.filedialog", fd),
                      ("tkinter.messagebox", mb)):
        monkeypatch.setitem(sys.modules, name, mod)
    _purge_gui_modules()
    common = importlib.import_module("dissertation.gui.common")
    entry = ttk.Entry(None)
    ev = type("E", (), {"widget": entry})()
    common._entry_focus(ev)
    for chunk in ("Глава", " один"):
        entry.insert("end", chunk)
        common._entry_record(ev)
    def key(sym, code, state=0):
        return type("E", (), {"keysym": sym, "keycode": code, "widget": entry, "state": state})()
    assert common._entry_ctrl(key("z", 90)) == "break"             # латиница
    assert entry.value == "Глава"
    assert common._entry_ctrl(key("Cyrillic_ya", 90)) in ("break", None)   # рус. (VK только на Windows)
    monkeypatch.setattr(common, "IS_WIN", True)
    assert common._entry_ctrl(key("Cyrillic_ya", 90)) == "break"   # ещё шаг назад
    assert entry.value == ""
    common._entry_ctrl(key("Cyrillic_en", 89))                      # Ctrl+Y — вернуть
    assert entry.value == "Глава"
    common._entry_ctrl(key("z", 90, state=1))                        # Ctrl+Shift+Z — вернуть
    assert entry.value == "Глава один"
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

    # --- выбор пункта показывает его содержимое; правка и сохранение
    iid = _find_iid(ed.tree, "1.1 Методы")
    ed.tree.selection_set(iid)
    ed._on_tree_select()
    assert ed.block() == B.SUBCHAPTER and ed.editing_el is not None
    assert tp.heading_var.get() == "Методы"
    assert tp.text.get("1.0", "end-1c") == "Текст подглавы."
    tp.heading_var.set("Методы исследования")
    tp.text.delete("1.0", "end")
    tp.text.insert("end", "Новый текст подглавы со ссылкой [@R1].")
    assert "1.1 Методы исследования" in preview_text()
    assert "вместо текущего содержимого" in ed.preview.header.get()
    ed.on_save_edit()
    texts = [p.text for p in ctl.builder.document.paragraphs]
    assert "1.1 Методы исследования" in texts and "Новый текст подглавы со ссылкой [2]." in texts  # R1 после сортировки — второй

    # --- «Новый блок» + вставка ещё одной подглавы, перемещение и удаление
    ed.on_new_block()
    assert ed.editing_el is None and tp.heading_var.get() == ""
    ed.pos_combo.current(3)                 # в конец выбранного раздела
    iid = _find_iid(ed.tree, "1 Обзор литературы")
    ed.tree.selection_set(iid)
    ed._on_tree_select()                    # загрузилась глава
    ed.on_new_block()
    ed.block_combo.current(ed._blocks.index(B.SUBCHAPTER))
    ed._on_block_change(user=True)
    tp.heading_var.set("Результаты")
    tp.text.insert("end", "Текст результатов.")
    ed.on_action()
    iid = _find_iid(ed.tree, "1.2 Результаты")
    assert iid
    ed.tree.selection_set(iid)
    ed._on_tree_select()
    ed.on_move(-1)
    assert _find_iid(ed.tree, "1.1 Результаты") and _find_iid(ed.tree, "1.2 Методы исследования")
    ed.on_move(-1)                          # выше некуда — сообщение, без ошибки
    assert "некуда" in ed.status_var.get()
    ed.on_delete()
    assert not _find_iid(ed.tree, "1.1 Результаты")
    assert _find_iid(ed.tree, "1.1 Методы исследования")

    # --- «＋ Добавить…»: пункт рядом с выбранным, уровень ◀/▶, замок
    from dissertation.outline import Kind

    iid = _find_iid(ed.tree, "1.1 Методы исследования")
    ed.tree.selection_set(iid)
    ed._on_tree_select()
    ed.on_add_menu()                                   # меню строится без ошибок
    ed.on_add_element(Kind.SECTION)
    assert _find_iid(ed.tree, "1.1.1 Новый пункт")
    assert ed.editing_el is not None and tp.heading_var.get() == "Новый пункт"
    tp.heading_var.set("Источники данных")
    ed.on_save_edit()
    assert _find_iid(ed.tree, "1.1.1 Источники данных")
    ed.on_level(-1)                                    # ◀ пункт → подглава
    assert _find_iid(ed.tree, "1.2 Источники данных")
    ed.on_level(1)                                     # ▶ обратно
    assert _find_iid(ed.tree, "1.1.1 Источники данных")
    ed.on_toggle_lock()
    iid = _find_iid(ed.tree, "1.1.1 Источники данных")
    assert ed.tree.item(iid, "values")[0] == "закр."
    assert ed.lock_btn.cget("text") == "Открепить"
    ed.on_level(-1)
    assert "закреплён" in ed.status_var.get()
    ed.on_toggle_lock()

    # --- приложения списком + ссылка {@P…}
    ed.on_add_element(Kind.APPENDIX)
    ap = ed.panels[B.APPENDIX]
    assert ed.current is ap and ap.selected_el is not None
    ap.heading_var.set("Исходные данные")
    ap.text.delete("1.0", "end")
    ap.text.insert("end", "Таблица значений.")
    ed.on_action()                                     # «Сохранить приложение»
    ap.on_new()
    ap.heading_var.set("Листинг программы")
    ap.text.insert("end", "Код.")
    assert "ПРИЛОЖЕНИЕ Б" in preview_text()
    ed.on_action()                                     # «Добавить приложение»
    assert [ap.tree.item(i, "values")[0] for i in ap.tree.get_children()] == ["А", "Б"]
    select(B.BODY)
    tp.text.insert("end", "Код — в приложении {@P2}.")
    assert "в приложении Б" in preview_text()
    ed.on_action()
    select(B.APPENDIX)
    ap.on_sort()
    assert ap.tree.item(ap.tree.get_children()[0], "values")[2] == "Листинг программы"
    assert "Код — в приложении А." in [p.text for p in ctl.builder.document.paragraphs]

    out = ctl.save(tmp_path / "gui.docx")
    texts = [p.text for p in Document(str(out)).paragraphs]
    assert texts.index("1.1 Методы исследования") < texts.index("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    assert not errors, errors
