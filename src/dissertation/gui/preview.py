"""Живой предпросмотр: приблизительный вид абзацев по профилю (шрифт, отступы, выравнивание).

Tk не умеет выравнивание «по ширине», поэтому такие абзацы показываются по
левому краю — остальное (кегль, полужирный, абзацный отступ, интервалы,
разрывы страниц) соответствует стилям документа.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from ..profile import FormattingProfile, ParagraphSpec
from ..style_manager import SPEC_ATTR, Role

PX_PER_CM = 30.0
PX_PER_PT = 1.15


def _tag_options(spec: ParagraphSpec) -> dict:
    size_px = max(9, int(round(spec.font_size_pt * PX_PER_PT)))
    style = []
    if spec.bold:
        style.append("bold")
    if spec.italic:
        style.append("italic")
    return dict(
        font=(spec.font_name, -size_px, *style),
        justify={"center": "center", "right": "right"}.get(spec.alignment, "left"),
        lmargin1=int((spec.left_indent_cm + spec.first_line_indent_cm) * PX_PER_CM),
        lmargin2=int(spec.left_indent_cm * PX_PER_CM),
        spacing1=int(spec.space_before_pt * PX_PER_PT) + 2,
        spacing2=int(size_px * max(0.0, spec.line_spacing - 1.0)),
        spacing3=int(spec.space_after_pt * PX_PER_PT) + 2,
    )


class PreviewView:
    def __init__(self, parent, profile: FormattingProfile):
        self.frame = ttk.Frame(parent)
        self.header = tk.StringVar(value="Предпросмотр")
        ttk.Label(self.frame, textvariable=self.header, foreground="#555").pack(anchor="w")
        self.warn = tk.StringVar()
        ttk.Label(self.frame, textvariable=self.warn, foreground="#b35c00").pack(anchor="w")
        box = ttk.Frame(self.frame)
        box.pack(fill="both", expand=True)
        self.text = tk.Text(box, wrap="word", height=8, relief="solid", borderwidth=1,
                            background="white", padx=24, pady=10, cursor="arrow")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set, state="disabled")
        self.text.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.set_profile(profile)

    def set_profile(self, profile: FormattingProfile) -> None:
        self.profile = profile
        self._caps = {}
        for role in Role:
            spec = getattr(profile, SPEC_ATTR.get(role, role.value))
            self.text.tag_configure(role.value, **_tag_options(spec))
            self._caps[role] = spec.all_caps
        base = profile.title_page
        for align in ("left", "center", "right"):
            for bold in (False, True):
                spec = ParagraphSpec(font_name=base.font_name, font_size_pt=base.font_size_pt,
                                     bold=bold, alignment=align, line_spacing=1.0,
                                     first_line_indent_cm=0)
                self.text.tag_configure(f"title_{align}_{int(bold)}", **_tag_options(spec))
        self.text.tag_configure("pagebreak", foreground="#999", justify="center",
                                font=("TkDefaultFont", 8), spacing1=6, spacing3=6)
        self.text.tag_configure("muted", foreground="#888", font=("TkDefaultFont", 9, "italic"))

    def _begin(self):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")

    def _end(self):
        self.text.configure(state="disabled")
        self.text.see("1.0")

    def show_message(self, message: str, header: str = "Предпросмотр") -> None:
        self.header.set(header)
        self.warn.set("")
        self._begin()
        self.text.insert("end", message, "muted")
        self._end()

    def show_result(self, result) -> None:
        """result: PreviewResult из DocumentBuilder.preview_block."""
        self.header.set("Предпросмотр" + (f" · будет вставлено {result.where}" if result.where else ""))
        self.warn.set(" · ".join(result.warnings))
        self._begin()
        if not result.items:
            self.text.insert("end", "Введите текст — здесь появится, как он будет оформлен.", "muted")
        breaks = result.page_breaks or [False] * len(result.items)
        for (role, value), page_break in zip(result.items, breaks):
            if page_break:
                self.text.insert("end", "─────────── новая страница ───────────\n", "pagebreak")
            shown = value.upper() if self._caps.get(role) else value
            self.text.insert("end", shown + "\n", role.value)
        self._end()

    def show_title(self, lines) -> None:
        """lines: title_page.preview_lines(...)."""
        self.header.set("Предпросмотр титульного листа (схематично, без масштаба высоты)")
        self.warn.set("")
        self._begin()
        for text, align, bold, gap in lines:
            if gap:
                self.text.insert("end", "\n" * gap)
            self.text.insert("end", (text or " ") + "\n", f"title_{align}_{int(bold)}")
        self._end()
