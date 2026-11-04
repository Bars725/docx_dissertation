"""Минимальная фейковая реализация tkinter для тестов GUI без дисплея/без Tk.

Повторяет только то, что использует наш интерфейс; неизвестные методы — no-op.
Проверяет логику обработчиков, связи с контроллером и предпросмотр; сам рендер Tk
проверяется тем же тестом на настоящем Tk (если он установлен).
"""
from __future__ import annotations

import itertools
import types

_ids = itertools.count(1)


class TclError(Exception):
    pass


class _Widget:
    def __init__(self, master=None, *args, **kw):
        self.master = master
        self._opts = dict(kw)
        self._bindings = {}

    def __getattr__(self, name):
        if name.startswith("__"):
            raise AttributeError(name)
        return lambda *a, **k: None

    # общие
    def configure(self, cnf=None, **kw):
        if isinstance(cnf, dict):
            kw = {**cnf, **kw}
        self._opts.update(kw)

    def config(self, cnf=None, **kw):
        return self.configure(cnf, **kw)

    def cget(self, key):
        return self._opts.get(key, "")

    def bind(self, seq, fn=None, add=None):
        self._bindings[seq] = fn

    def fire(self, seq, event=None):
        fn = self._bindings.get(seq)
        return fn(event) if fn else None

    def winfo_children(self):
        return []

    def after(self, ms, fn=None, *args):
        if fn:
            fn(*args)
        return f"after#{next(_ids)}"

    def after_cancel(self, _id):
        pass


class Tk(_Widget):
    def __init__(self, *a, **k):
        super().__init__(None)
        self.clipboard = ""

    def clipboard_get(self):
        return self.clipboard


class Toplevel(_Widget):
    pass


class Menu(_Widget):
    pass


class Canvas(_Widget):
    def bbox(self, *a):
        return (0, 0, 0, 0)

    def create_window(self, *a, **k):
        return 1


class Listbox(_Widget):
    def __init__(self, master=None, **kw):
        super().__init__(master, **kw)
        self.items, self.sel = [], []

    def insert(self, index, *items):
        self.items.extend(items)

    def delete(self, first, last=None):
        self.items = []

    def curselection(self):
        return tuple(self.sel)

    def selection_set(self, i):
        self.sel = [i]


class Text(_Widget):
    def __init__(self, master=None, **kw):
        super().__init__(master, **kw)
        self.content = ""
        self.tagged = []          # [(text, tags)]
        self.modified = False
        self.tags = {}

    def _touch(self):
        was = self.modified
        self.modified = True
        if not was:
            self.fire("<<Modified>>")

    def insert(self, index, chars, *tags):
        if self.cget("state") == "disabled":
            return
        if index == "1.0":
            self.content = chars + self.content
        else:
            self.content += chars
        self.tagged.append((chars, tags))
        self._touch()

    def delete(self, first, last=None):
        if self.cget("state") == "disabled":
            return
        if first.startswith("sel"):
            raise TclError("no selection")
        self.content = ""
        self.tagged = []
        self._touch()

    def get(self, first="1.0", last="end"):
        return self.content if last == "end-1c" else self.content + "\n"

    def edit_modified(self, value=None):
        if value is None:
            return self.modified
        self.modified = bool(value)

    def tag_configure(self, name, **kw):
        self.tags[name] = kw


class Variable:
    _default = ""

    def __init__(self, master=None, value=None):
        self._value = self._default if value is None else value
        self._traces = []

    def get(self):
        return self._value

    def set(self, value):
        self._value = value
        for cb in self._traces:
            cb("var", "", "write")

    def trace_add(self, mode, cb):
        self._traces.append(cb)


class StringVar(Variable):
    pass


class BooleanVar(Variable):
    _default = False


# ------------------------------------------------------------------ ttk
class Frame(_Widget):
    pass


class Label(_Widget):
    pass


class Button(_Widget):
    def invoke(self):
        return self._opts["command"]()


class Entry(_Widget):
    def __init__(self, master=None, **kw):
        super().__init__(master, **kw)
        self.value = ""

    def insert(self, index, s):
        self.value += s

    def delete(self, first, last=None):
        if str(first).startswith("sel"):
            raise TclError("no selection")
        self.value = ""

    def clipboard_get(self):
        return self._opts.get("clipboard", "")


class Checkbutton(_Widget):
    pass


class Radiobutton(_Widget):
    pass


class Scrollbar(_Widget):
    pass


class PanedWindow(_Widget):
    pass


class Style(_Widget):
    def configure(self, style=None, **kw):
        pass

    def theme_names(self):
        return ("clam",)


class Combobox(_Widget):
    def __init__(self, master=None, **kw):
        super().__init__(master, **kw)
        self._cur = 0

    def current(self, index=None):
        if index is None:
            return self._cur
        self._cur = index

    def select(self, index):
        self._cur = index
        self.fire("<<ComboboxSelected>>")


class Treeview(_Widget):
    def __init__(self, master=None, **kw):
        super().__init__(master, **kw)
        self.rows = {}         # iid -> dict(parent, text, values)
        self.order = []
        self._sel = ()

    def insert(self, parent, index, text="", values=(), **kw):
        iid = f"I{next(_ids)}"
        self.rows[iid] = dict(parent=parent, text=text, values=tuple(values))
        self.order.append(iid)
        return iid

    def delete(self, *iids):
        for iid in iids:
            self.rows.pop(iid, None)
            if iid in self.order:
                self.order.remove(iid)
        self._sel = tuple(i for i in self._sel if i in self.rows)

    def get_children(self, item=""):
        return tuple(i for i in self.order if self.rows[i]["parent"] == item)

    def selection(self):
        return self._sel

    def selection_set(self, *iids):
        self._sel = tuple(iids)

    def selection_remove(self, *iids):
        self._sel = tuple(i for i in self._sel if i not in iids)

    def item(self, iid, option=None):
        row = self.rows[iid]
        return row[option] if option else row

    def select(self, iid):
        self.selection_set(iid)
        self.fire("<<TreeviewSelect>>")


def build_modules():
    """(tkinter, ttk, filedialog, messagebox) — модули-заглушки."""
    tk = types.ModuleType("tkinter")
    for name in ("Tk", "Toplevel", "Menu", "Canvas", "Listbox", "Text", "StringVar", "BooleanVar"):
        setattr(tk, name, globals()[name])
    tk.TclError = TclError
    tk.Misc = _Widget
    ttk = types.ModuleType("tkinter.ttk")
    for name in ("Frame", "Label", "Button", "Entry", "Checkbutton", "Radiobutton", "Scrollbar",
                 "PanedWindow", "Style", "Combobox", "Treeview"):
        setattr(ttk, name, globals()[name])
    fd = types.ModuleType("tkinter.filedialog")
    fd.askopenfilename = lambda **k: ""
    fd.asksaveasfilename = lambda **k: ""
    mb = types.ModuleType("tkinter.messagebox")
    mb.log = []

    def showerror(title, message, **k):
        mb.log.append(("error", message))

    mb.showerror = showerror
    mb.showinfo = lambda title, message, **k: mb.log.append(("info", message))
    mb.askyesno = lambda *a, **k: True
    mb.askyesnocancel = lambda *a, **k: False
    tk.ttk, tk.filedialog, tk.messagebox = ttk, fd, mb
    return tk, ttk, fd, mb
