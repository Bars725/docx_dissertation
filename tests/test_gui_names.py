"""Наши классы GUI не должны перекрывать атрибуты базовых виджетов Tk.

Пример бага: EditorFrame._options() перекрыл внутренний Misc._options(cnf, kw),
и создание любого дочернего виджета падало с TypeError. Новые части GUI
написаны композицией (виджет — в атрибуте), подклассы Tk остались только у диалогов.
"""
import inspect
import sys

import pytest

import fake_tk

ALLOWED = {"__init__", "__doc__", "__module__", "__qualname__", "__dict__", "__weakref__",
           "__firstlineno__", "__static_attributes__"}

# Имена методов/атрибутов tkinter.Misc / BaseWidget / Wm / Pack / Grid / Place / ttk.Widget,
# которые легко перекрыть случайно (для проверки без установленного Tk).
KNOWN_TK_NAMES = set("""
_options _register _root _nametowidget _getints _getdoubles _getboolean _displayof
_windowingsystem _do _configure _getconfigure _getconfigure1 _report_exception _substitute
_subst_format _subst_format_str _bind _unbind _grid_configure _gridconvvalue _setup
_last_child_ids _w _name _tclCommands master tk children widgetName
after after_idle after_cancel bell bind bind_all bind_class bindtags cget clipboard_append
clipboard_clear clipboard_get config configure deletecommand destroy event_add event_delete
event_generate event_info focus focus_displayof focus_force focus_get focus_lastfor focus_set
getboolean getdouble getint getvar grab_current grab_release grab_set grab_set_global
grab_status image_names image_types info_patchlevel keys lift lower mainloop nametowidget
option_add option_clear option_get option_readfile quit register selection_clear
selection_get selection_handle selection_own selection_own_get send setvar tk_bisque
tk_focusFollowsMouse tk_focusNext tk_focusPrev tk_setPalette tk_strictMotif tkraise unbind
unbind_all unbind_class update update_idletasks wait_variable wait_visibility wait_window
waitvar winfo_children winfo_exists winfo_height winfo_width winfo_toplevel
aspect attributes client colormapwindows command deiconify focusmodel forget frame geometry
grid group iconbitmap iconify iconmask iconname iconphoto iconposition iconwindow manage
maxsize minsize overrideredirect positionfrom protocol resizable sizefrom state title
transient withdraw pack place info slaves propagate location size bbox columnconfigure
rowconfigure anchor content identify instate
""".split())


def _gui_classes(gui_modules):
    for mod in gui_modules:
        for _, cls in inspect.getmembers(mod, inspect.isclass):
            if cls.__module__ == mod.__name__:
                yield cls


def test_no_known_tk_names_overridden(monkeypatch):
    """Работает без Tk: фейковые модули + статический список имён."""
    tk, ttk, fd, mb = fake_tk.build_modules()
    for name, mod in (("tkinter", tk), ("tkinter.ttk", ttk), ("tkinter.filedialog", fd),
                      ("tkinter.messagebox", mb)):
        monkeypatch.setitem(sys.modules, name, mod)
    for name in [m for m in sys.modules if m.startswith("dissertation.gui")]:
        monkeypatch.delitem(sys.modules, name)
    import importlib

    mods = [importlib.import_module(f"dissertation.gui.{m}")
            for m in ("app", "dialogs", "editor", "panels", "preview", "common")]
    bad = {}
    for cls in _gui_classes(mods):
        widget_like = any(b.__module__.startswith(("tkinter", "fake_tk")) for b in cls.__mro__[1:])
        if widget_like:
            clashes = (set(vars(cls)) - ALLOWED) & KNOWN_TK_NAMES
            if clashes:
                bad[cls.__name__] = sorted(clashes)
    for name in [m for m in sys.modules if m.startswith("dissertation.gui")]:
        del sys.modules[name]
    assert not bad, f"перекрыты атрибуты Tk: {bad}"


def test_no_overridden_tk_attributes_real():
    tkinter = pytest.importorskip("tkinter")
    for name in [m for m in sys.modules if m.startswith("dissertation.gui")]:
        del sys.modules[name]
    import importlib

    mods = [importlib.import_module(f"dissertation.gui.{m}")
            for m in ("app", "dialogs", "editor", "panels", "preview", "common")]
    bad = {}
    for cls in _gui_classes(mods):
        if not issubclass(cls, tkinter.Misc):
            continue
        tk_bases = [b for b in cls.__mro__[1:] if b.__module__.startswith("tkinter")]
        tk_names = set().union(*(dir(b) for b in tk_bases))
        clashes = (set(vars(cls)) - ALLOWED) & tk_names
        if clashes:
            bad[cls.__name__] = sorted(clashes)
    assert not bad, f"перекрыты атрибуты Tk: {bad}"
