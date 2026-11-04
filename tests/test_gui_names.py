"""Наши классы GUI не должны перекрывать атрибуты базовых виджетов Tk.

Пример бага: EditorFrame._options() перекрыл внутренний Misc._options(cnf, kw),
и создание любого дочернего виджета падало с TypeError.
"""
import inspect

import pytest

tkinter = pytest.importorskip("tkinter")

from dissertation.gui import app as gui  # noqa: E402

ALLOWED = {"__init__", "__doc__", "__module__", "__qualname__", "__dict__", "__weakref__"}


@pytest.mark.parametrize("cls", [c for _, c in inspect.getmembers(gui, inspect.isclass)
                                 if issubclass(c, tkinter.Misc) and c.__module__ == gui.__name__])
def test_no_overridden_tk_attributes(cls):
    tk_bases = [b for b in cls.__mro__[1:] if b.__module__.startswith("tkinter")]
    tk_names = set().union(*(dir(b) for b in tk_bases))
    clashes = (set(vars(cls)) - ALLOWED) & tk_names
    assert not clashes, f"{cls.__name__} перекрывает атрибуты Tk: {sorted(clashes)}"
