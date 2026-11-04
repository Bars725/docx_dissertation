"""Операции над структурой: чтение/правка содержимого пункта, перемещение, удаление.

«Раздел» пункта — его заголовок, собственный текст и все вложенные подпункты
(до следующего заголовка того же или более высокого уровня). «Собственный текст» —
абзацы между заголовком и первым следующим заголовком любого уровня.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from . import fields, ooxml
from .bibliography import BM_PREFIX
from .xref import APP_PREFIX, FIG_PREFIX, XRef, has_refs
from .outline import Kind, OutlineItem, StyleResolver
from .runs import make_run, set_paragraph_text, visible_text
from .style_manager import Role
from .text_cleaner import CleanOptions, TextCleaner

MOVABLE = {Kind.CHAPTER, Kind.SUBCHAPTER, Kind.SECTION, Kind.APPENDIX, Kind.OTHER_STRUCTURAL}
LEVEL_KINDS = [Kind.CHAPTER, Kind.SUBCHAPTER, Kind.SECTION]
LOCK_PREFIX = "_GostLock_"
KIND_RU = {Kind.CHAPTER: "глава", Kind.SUBCHAPTER: "подглава", Kind.SECTION: "пункт",
           Kind.APPENDIX: "приложение", Kind.INTRODUCTION: "введение", Kind.CONCLUSION: "заключение",
           Kind.BIBLIOGRAPHY: "список литературы", Kind.TOC: "оглавление",
           Kind.TITLE: "титульный лист", Kind.OTHER_STRUCTURAL: "структурный элемент"}
_REF_MARK = re.compile(r"\x01([A-Za-z][A-Za-z0-9]*)\x02")
# Сохраняемый элемент в тексте пункта: ⟦#3 Таблица⟧, ⟦#1 Рисунок 2 – Схема⟧
KEEP_RX = re.compile(r"⟦#(\d+)(?:\s+([^⟧]*))?⟧")
# Эти элементы можно «изменить принудительно» (превратить в обычный текст)
SOFT = {"список", "сноска", "поле Word"}
_INVISIBLE = {qn("w:bookmarkStart"), qn("w:bookmarkEnd"), qn("w:proofErr"), qn("w:permStart"),
              qn("w:permEnd"), qn("w:commentRangeStart"), qn("w:commentRangeEnd")}


@dataclass
class SectionContent:
    kind: Kind
    heading: str = ""
    heading_editable: bool = True
    body: str = ""
    body_editable: bool = True
    reason: str = ""                     # пояснение (маркеры, защищённые элементы)
    paragraphs: int = 0
    soft_protected: int = 0              # абзацы, которые можно «изменить принудительно»
    force: bool = False


@dataclass
class OpResult:
    ok: bool
    message: str
    warnings: list[str] = field(default_factory=list)
    element: object = None          # созданный/затронутый элемент (для выделения в GUI)


class StructureOps:
    def __init__(self, builder):
        self.b = builder
        self.doc = builder.document
        self.body_el = builder.document.element.body

    # ========================================================= навигация
    def _item(self, element) -> tuple[int, list[OutlineItem]]:
        outline = self.b.outline()
        for i, it in enumerate(outline):
            if it.element is element:
                return i, outline
        raise ValueError("Элемент структуры не найден — обновите структуру")

    def _top(self, el):
        while el is not None and el.getparent() is not self.body_el:
            el = el.getparent()
        return el

    def _children(self) -> list:
        return [c for c in self.body_el if c.tag != ooxml.W_SECTPR]

    def section_elements(self, element) -> list:
        """Заголовок + собственный текст + подпункты (элементы верхнего уровня тела)."""
        i, outline = self._item(element)
        item = outline[i]
        if item.kind == Kind.TITLE:
            return [self._top(it.element) for it in outline if it.kind == Kind.TITLE]
        children = self._children()
        start = children.index(self._top(item.element))
        end = len(children)
        for later in outline[i + 1:]:
            if later.kind == Kind.TITLE:
                continue
            if later.level <= item.level:
                end = children.index(self._top(later.element))
                break
        return children[start:end]

    def own_body(self, element) -> list:
        """Собственные абзацы пункта (без заголовка, заголовка приложения и подпунктов)."""
        i, outline = self._item(element)
        children = self._children()
        start = children.index(self._top(outline[i].element)) + 1
        end = len(children)
        for later in outline[i + 1:]:
            end = children.index(self._top(later.element))
            break
        els = children[start:end]
        resolver = StyleResolver(self.doc)
        if outline[i].kind == Kind.APPENDIX and els and els[0].tag == ooxml.W_P \
                and resolver.role_of(els[0]) == Role.APPENDIX_TITLE:
            els = els[1:]
        return els

    def _appendix_title(self, element):
        nxt = self._top(element).getnext()
        if nxt is not None and nxt.tag == ooxml.W_P and \
                StyleResolver(self.doc).role_of(nxt) == Role.APPENDIX_TITLE:
            return nxt
        return None

    def parent_and_siblings(self, element):
        i, outline = self._item(element)
        item = outline[i]

        def parent_of(j):
            for k in range(j - 1, -1, -1):
                if outline[k].level < outline[j].level and outline[k].kind != Kind.TITLE:
                    return k
            return None

        p = parent_of(i)
        sibs = [j for j, it in enumerate(outline)
                if it.level == item.level and it.kind != Kind.TITLE and parent_of(j) == p]
        return outline, i, sibs

    # ========================================================= чтение
    # ========================================================= группы
    def _groups(self, element, force: bool = False) -> list[dict]:
        """Собственный текст пункта как последовательность групп:
        text — обычный абзац; keep — сохраняемый элемент(ы) с подписью-маркером;
        marker — невидимые служебные элементы (закладки) — не трогаем."""
        from .figures import Figures

        resolver = StyleResolver(self.doc)
        figs = {id(e.caption): e for e in Figures(self.b).entries()}
        groups, skip = [], set()
        els = self.own_body(element)
        for idx, el in enumerate(els):
            if id(el) in skip:
                continue
            if el.tag in _INVISIBLE:
                groups.append({"kind": "marker", "els": [el]})
                continue
            if el.tag == ooxml.W_P and resolver.role_of(el) == Role.FIGURE:
                nxt = els[idx + 1] if idx + 1 < len(els) else None
                if nxt is not None and id(nxt) in figs:
                    e = figs[id(nxt)]
                    skip.add(id(nxt))
                    groups.append({"kind": "keep", "els": [el, nxt], "figure": e,
                                   "label": self.b.profile.figure_caption_format.format(
                                       n=e.number, title=e.title)})
                    continue
            if el.tag == ooxml.W_P and id(el) in figs:          # подпись без рисунка
                e = figs[id(el)]
                groups.append({"kind": "keep", "els": [el], "figure": e,
                               "label": self.b.profile.figure_caption_format.format(n=e.number, title=e.title)})
                continue
            reason = self._complex(el)
            if reason is None or (force and reason in SOFT):
                if el.tag == ooxml.W_P:
                    groups.append({"kind": "text", "els": [el], "text": self._editable_text(el),
                                   "soft": reason})
                continue
            preview = " ".join(ooxml.element_text(el).split())
            label = reason.capitalize() + (f": {preview[:40]}…" if len(preview) > 40 else
                                           (f": {preview}" if preview else ""))
            groups.append({"kind": "keep", "els": [el], "label": label, "soft": reason in SOFT})
        n = 0
        for g in groups:
            if g["kind"] == "keep":
                n += 1
                g["n"] = n
        return groups

    # ========================================================= чтение
    def read(self, element, force: bool = False) -> SectionContent:
        i, outline = self._item(element)
        item = outline[i]
        content = SectionContent(kind=item.kind, force=force)
        para = Paragraph(item.element, self.doc._body) if item.element.tag == ooxml.W_P else None
        if item.kind in (Kind.CHAPTER, Kind.SUBCHAPTER, Kind.SECTION):
            content.heading = self._strip_number(item)
        elif item.kind == Kind.APPENDIX:
            t = self._appendix_title(item.element)
            content.heading = visible_text(Paragraph(t, self.doc._body)) if t is not None else ""
        else:
            content.heading = visible_text(para) if para is not None else item.text
            content.heading_editable = False
        parts, kept = [], []
        for g in self._groups(element, force):
            if g["kind"] == "text" and g["text"].strip():
                parts.append(g["text"])
                content.paragraphs += 1
            elif g["kind"] == "keep":
                parts.append(f"⟦#{g['n']} {g['label']}⟧")
                kept.append(g["label"].split(":")[0])
                if g.get("soft"):
                    content.soft_protected += 1
        content.body = "\n\n".join(parts)
        if kept:
            content.reason = ("сложные элементы показаны маркерами ⟦#…⟧ — их можно переставлять "
                              "или удалить строкой целиком; " + ", ".join(sorted(set(kept))).lower())
        return content

    def _strip_number(self, item: OutlineItem) -> str:
        from .document_builder import _format_regex

        num = self.b.profile.numbering
        fmt = {Kind.CHAPTER: num.chapter, Kind.SUBCHAPTER: num.subchapter,
               Kind.SECTION: num.section}[item.kind]
        text = " ".join(ooxml.element_text(item.element).split())
        if num.enabled:
            m = _format_regex(fmt).match(text)
            if m:
                return m.group("title").strip()
        return text

    @staticmethod
    def _complex(el) -> str | None:
        if el.tag == ooxml.W_TBL:
            return "таблица"
        if el.tag != ooxml.W_P:
            return "встроенный элемент"
        for tag, name in (("w:drawing", "рисунок"), ("w:pict", "рисунок"), ("w:object", "объект"),
                          ("w:footnoteReference", "сноска"), ("w:endnoteReference", "сноска"),
                          ("w:txbxContent", "надпись")):
            if el.find(".//" + qn(tag)) is not None:
                return name
        if el.find(".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath") is not None:
            return "формула"
        ppr = el.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:numPr")) is not None:
            return "список"
        for f in fields.iter_complex_fields(el):
            target = fields.ref_target(f.instr) or ""
            if not target.startswith((BM_PREFIX, APP_PREFIX, FIG_PREFIX)):
                return "поле Word"
        return None

    def _editable_text(self, p_el) -> str:
        """Текст абзаца; поля-ссылки на литературу → @КЛЮЧ (получается [@R3, с. 15])."""
        out = []
        stack = []
        for r in p_el.iter(qn("w:r")):
            for child in r:
                if child.tag == fields.W_FLDCHAR:
                    kind = child.get(qn("w:fldCharType"))
                    if kind == "begin":
                        stack.append({"instr": "", "state": "instr"})
                    elif kind == "separate" and stack:
                        stack[-1]["state"] = "result"
                    elif kind == "end" and stack:
                        f = stack.pop()
                        target = fields.ref_target(f["instr"]) or ""
                        if target.startswith(BM_PREFIX):
                            out.append(f"\x01{target[len(BM_PREFIX):]}\x02")
                        elif target.startswith((APP_PREFIX, FIG_PREFIX)):
                            out.append("{@" + target.split("_")[-1] + "}")
                elif child.tag == fields.W_INSTR and stack:
                    stack[-1]["instr"] += child.text or ""
                elif stack:
                    continue          # результат поля пропускаем — восстановим токеном
                elif child.tag == qn("w:t"):
                    out.append(child.text or "")
                elif child.tag == qn("w:tab"):
                    out.append("\t")
                elif child.tag in (qn("w:br"), qn("w:cr")):
                    out.append("\n")
        return _REF_MARK.sub(lambda m: "@" + m.group(1), "".join(out))

    # ========================================================= правка
    # ========================================================= правка
    def _segments(self, body: str, clean: bool, cleaner) -> list:
        """Текст из окна → [('text', абзац) | ('keep', n, подпись) | ('figure', FigureSpec)]."""
        from .figures import FigureSpec

        out, pos = [], 0
        for m in KEEP_RX.finditer(body):
            if m.start() > pos:
                out += self._text_segments(body[pos:m.start()], clean, cleaner)
            out.append(("keep", int(m.group(1)), (m.group(2) or "").strip()))
            pos = m.end()
        out += self._text_segments(body[pos:], clean, cleaner)
        return [x if not isinstance(x, tuple) or x[0] != "fig" else ("figure", x[1]) for x in out]

    def _text_segments(self, text: str, clean: bool, cleaner) -> list:
        items = self.b._body_items(text, clean, cleaner)
        return [("text", v) if r == Role.BODY else ("figure", v) for r, v in items]

    def update(self, element, heading: str | None, body: str | None, clean: bool = True,
               options: CleanOptions | None = None, force: bool = False) -> OpResult:
        content = self.read(element, force)
        i, outline = self._item(element)
        item = outline[i]
        cleaner = TextCleaner(options or CleanOptions())
        warnings: list[str] = []
        if heading is not None and content.heading_editable:
            title = cleaner.heading(heading)
            if not title:
                return OpResult(False, "Заголовок не может быть пустым")
            if item.kind == Kind.APPENDIX:
                t = self._appendix_title(item.element)
                if t is not None:
                    set_paragraph_text(Paragraph(t, self.doc._body), title)
                else:
                    add = self.b._inserter("after", self._top(item.element))
                    add(Role.APPENDIX_TITLE, title)
            else:
                from .document_builder import BlockType

                block = {Kind.CHAPTER: BlockType.CHAPTER, Kind.SUBCHAPTER: BlockType.SUBCHAPTER,
                         Kind.SECTION: BlockType.SECTION}[item.kind]
                set_paragraph_text(Paragraph(item.element, self.doc._body),
                                   self.b._numbered_heading(block, title))
        if body is not None:
            warnings += self._replace_body(element, item, body, clean, cleaner, force)
        self.b.renumber()
        return OpResult(True, "Изменения сохранены", warnings)

    def _replace_body(self, element, item, body, clean, cleaner, force) -> list[str]:
        from .figures import Figures, insert_figure_after

        groups = self._groups(element, force)
        keeps = {g["n"]: g for g in groups if g["kind"] == "keep"}
        segments = self._segments(body, clean, cleaner)
        used: set[int] = set()
        warnings = []
        # убрать старое (кроме служебных закладок), сохраняемые элементы — отложить
        for g in groups:
            if g["kind"] == "marker":
                continue
            for el in g["els"]:
                el.getparent().remove(el)
        cursor = self._appendix_title(item.element) if item.kind == Kind.APPENDIX else None
        cursor = cursor if cursor is not None else self._top(item.element)
        style = self.b.styles.style_for(Role.BODY)
        self.b._unknown_refs = []
        from .xref import XRef

        xr = XRef(self.b)
        for seg in segments:
            if seg[0] == "text":
                p = ooxml.new_paragraph_after(cursor, self.doc._body)
                p.style = style
                text = seg[1]
                if has_refs(text):
                    p._p.extend(xr.rich_runs(text, self.b._unknown_refs))
                else:
                    p._p.append(make_run(text))
                cursor = p._p
            elif seg[0] == "figure":
                cursor = insert_figure_after(self.b, cursor, seg[1])
            else:
                n, label = seg[1], seg[2]
                g = keeps.get(n)
                if g is None or n in used:
                    warnings.append(f"Маркер ⟦#{n}⟧ не найден — пропущен")
                    continue
                used.add(n)
                for el in g["els"]:
                    cursor.addnext(el)
                    cursor = el
                fig = g.get("figure")
                if fig is not None and label:
                    m = re.match(r".*?\s[–—-]\s(.+)$", label)
                    new_title = (m.group(1) if m else label).strip()
                    if new_title and new_title != fig.title:
                        Figures(self.b).set_title(fig.caption, new_title)
        dropped = [g["label"].split(":")[0] for n, g in keeps.items() if n not in used]
        if dropped:
            warnings.append("Удалены элементы: " + ", ".join(dropped))
        if self.b._unknown_refs:
            warnings.append("Неизвестные ссылки: " + ", ".join(sorted(set(self.b._unknown_refs))))
        return warnings

    def preview_update(self, element, heading: str, body: str, clean: bool = True,
                       options: CleanOptions | None = None, force: bool = False):
        """Предпросмотр правки: тот же номер пункта, новый заголовок и текст."""
        from .document_builder import PreviewResult

        content = self.read(element, force)
        i, outline = self._item(element)
        item = outline[i]
        cleaner = TextCleaner(options or CleanOptions())
        res = PreviewResult(where=f"вместо текущего содержимого «{item.text}»")
        role = {Kind.CHAPTER: Role.CHAPTER, Kind.SUBCHAPTER: Role.SUBCHAPTER,
                Kind.SECTION: Role.SECTION, Kind.APPENDIX: Role.APPENDIX}.get(item.kind, Role.STRUCTURAL)
        title = cleaner.heading(heading) if heading.strip() else ""
        if item.kind in (Kind.CHAPTER, Kind.SUBCHAPTER, Kind.SECTION) and content.heading_editable:
            full = " ".join(ooxml.element_text(item.element).split())
            old_title = content.heading
            prefix = full[: len(full) - len(old_title)] if full.endswith(old_title) else ""
            res.items.append((role, prefix + (title or old_title)))
        else:
            res.items.append((role, item.text))
            if item.kind == Kind.APPENDIX and (title or content.heading):
                res.items.append((Role.APPENDIX_TITLE, title or content.heading))
        xr = XRef(self.b)
        keeps = {g["n"]: g for g in self._groups(element, force) if g["kind"] == "keep"}
        figs_done = 0
        for seg in self._segments(body, clean, cleaner):
            if seg[0] == "text":
                res.items.append((Role.BODY, xr.render(seg[1])))
            elif seg[0] == "figure":
                res.items += self.b.figure_preview_items(seg[1], figs_done)
                figs_done += 1
            else:
                g = keeps.get(seg[1])
                res.items.append((Role.FIGURE_CAPTION if g and g.get("figure") else Role.BODY,
                                  f"[{seg[2] or 'элемент'}]" if g else f"[⟦#{seg[1]}⟧ не найден]"))
        if content.soft_protected and not force:
            res.warnings.append("Списки/сноски показаны маркерами — «Изменить принудительно» "
                                "превратит их в обычный текст")
        res.page_breaks = [self.b._role_breaks(r) for r, _ in res.items]
        return res

    # ======================================================= замок
    def is_locked(self, element) -> bool:
        return any((bm.get(qn("w:name")) or "").startswith(LOCK_PREFIX)
                   for bm in element.iter(fields.W_BM_START))

    def set_locked(self, element, locked: bool) -> OpResult:
        """Закрепить пункт на месте: невидимая закладка в заголовке (переживает Word/LO)."""
        if locked == self.is_locked(element):
            return OpResult(True, "Без изменений")
        if locked:
            from docx.oxml import OxmlElement

            bm_id = fields.next_bookmark_id(self.doc)
            start = OxmlElement("w:bookmarkStart")
            start.set(qn("w:id"), str(bm_id))
            start.set(qn("w:name"), f"{LOCK_PREFIX}{bm_id}")
            end = OxmlElement("w:bookmarkEnd")
            end.set(qn("w:id"), str(bm_id))
            ppr = element.find(qn("w:pPr"))
            if ppr is not None:
                ppr.addnext(start)
            else:
                element.insert(0, start)
            start.addnext(end)
            return OpResult(True, "Пункт закреплён: его место и уровень не изменятся")
        for bm in list(element.iter(fields.W_BM_START)):
            if (bm.get(qn("w:name")) or "").startswith(LOCK_PREFIX):
                bm_id = bm.get(qn("w:id"))
                for end in list(element.iter(fields.W_BM_END)):
                    if end.get(qn("w:id")) == bm_id:
                        end.getparent().remove(end)
                bm.getparent().remove(bm)
        return OpResult(True, "Пункт откреплён")

    def _locked_msg(self, item) -> OpResult:
        return OpResult(False, f"«{item.text}» закреплён — снимите замок, чтобы изменить его место")

    # ======================================================= перемещение
    def _move_block(self, elems, *, before=None, after=None) -> None:
        if before is not None:
            for el in elems:
                before.addprevious(el)
        else:
            for el in reversed(elems):
                after.addnext(el)

    def _own_end(self, element):
        """Последний элемент «собственного» текста пункта (или сам заголовок)."""
        body = self.own_body(element)
        if self._items_kind(element) == Kind.APPENDIX:
            t = self._appendix_title(element)
            if t is not None and not body:
                return t
        return body[-1] if body else self._top(element)

    def _items_kind(self, element) -> Kind:
        i, outline = self._item(element)
        return outline[i].kind

    def move(self, element, delta: int) -> OpResult:
        """Вверх/вниз. Глава ↔ глава, приложение ↔ приложение; подглава и пункт могут
        переходить в соседнюю главу/подглаву. Порядок разделов ГОСТ не нарушается."""
        i, outline = self._item(element)
        item = outline[i]
        if self.is_locked(item.element):
            return self._locked_msg(item)
        if item.kind not in MOVABLE:
            return OpResult(False, f"«{item.text}» — {KIND_RU.get(item.kind, 'элемент')}: его место "
                                   "определено порядком ГОСТ (титул, оглавление, введение, основная "
                                   "часть, заключение, список литературы, приложения)")
        up = delta < 0
        outline, i, sibs = self.parent_and_siblings(element)
        pos = sibs.index(i)
        j = pos - 1 if up else pos + 1
        if 0 <= j < len(sibs):
            neighbor = outline[sibs[j]]
            if neighbor.kind == item.kind:
                if self.is_locked(neighbor.element):
                    return OpResult(False, f"Соседний пункт «{neighbor.text}» закреплён")
                mine = self.section_elements(item.element)
                theirs = self.section_elements(neighbor.element)
                if up:
                    self._move_block(mine, before=theirs[0])
                else:
                    self._move_block(mine, after=theirs[-1])
                self.b.renumber()
                return OpResult(True, "Перемещено, номера пересчитаны")
            if item.level == 0:
                return OpResult(False, f"{KIND_RU[item.kind].capitalize()} не может стоять "
                                       f"{'выше' if up else 'ниже'} — там «{neighbor.text}» (порядок ГОСТ)")
        if item.level == 0:
            return OpResult(False, "Дальше перемещать некуда")
        # --- переход в соседнего «родителя»: подглава → в соседнюю главу, пункт → в соседнюю подглаву
        parent_kind = LEVEL_KINDS[item.level - 1]
        parents = [it for it in outline if it.kind == parent_kind]
        my_parent = next(outline[k] for k in range(i - 1, -1, -1) if outline[k].level < item.level)
        if my_parent.kind != parent_kind or my_parent not in parents:
            return OpResult(False, "Дальше перемещать некуда")
        pidx = parents.index(my_parent)
        target_idx = pidx - 1 if up else pidx + 1
        if not 0 <= target_idx < len(parents):
            return OpResult(False, "Дальше перемещать некуда")
        target = parents[target_idx]
        mine = self.section_elements(item.element)
        if up:       # стать последним в предыдущем родителе
            self._move_block(mine, after=self.section_elements(target.element)[-1])
        else:        # стать первым в следующем родителе
            self._move_block(mine, after=self._own_end(target.element))
        self.b.renumber()
        return OpResult(True, f"Перемещено в «{target.text}», номера пересчитаны")

    # ======================================================= уровень
    def change_level(self, element, delta: int) -> OpResult:
        """← повысить (пункт → подглава → глава), → понизить; вместе с подпунктами."""
        i, outline = self._item(element)
        item = outline[i]
        if item.kind not in LEVEL_KINDS:
            return OpResult(False, f"У элемента «{item.text}» нет уровня глава/подглава/пункт")
        if self.is_locked(item.element):
            return self._locked_msg(item)
        new_level = item.level + delta
        if new_level < 0:
            return OpResult(False, "Это уже глава — выше уровня нет")
        sub = [it for it in outline[i:] if it is item or (
            it.level > item.level and self._top(it.element) in self.section_elements(item.element))]
        if max(it.level for it in sub) + delta > 2:
            return OpResult(False, "Глубже пункта (3-й уровень) по ГОСТ не нумеруется — "
                                   "сначала поднимите вложенные пункты")
        if delta > 0:
            prev = next((outline[k] for k in range(i - 1, -1, -1) if outline[k].level <= item.level), None)
            if prev is None or prev.kind != item.kind or prev.level != item.level:
                return OpResult(False, "Понизить нельзя: выше нет "
                                       f"{'главы' if item.level == 0 else 'подглавы'}, в которую можно вложить")
        from .document_builder import BlockType

        own_title = self._strip_number(item)
        block_of = {0: BlockType.CHAPTER, 1: BlockType.SUBCHAPTER, 2: BlockType.SECTION}
        role_of = {0: Role.CHAPTER, 1: Role.SUBCHAPTER, 2: Role.SECTION}
        for it in sub:
            title = self._strip_number(it)
            lvl = it.level + delta
            para = Paragraph(it.element, self.doc._body)
            para.style = self.b.styles.style_for(role_of[lvl])
            set_paragraph_text(para, self.b._numbered_heading(block_of[lvl], title))
        self.b.renumber()
        what = {0: "главой", 1: "подглавой", 2: "пунктом"}[new_level]
        return OpResult(True, f"«{own_title}» теперь {what}, номера пересчитаны")

    # ======================================================= добавление
    def available_elements(self, near=None) -> list[tuple[Kind, str, bool, str]]:
        """[(вид, подпись, доступен, причина)] для меню «Добавить элемент структуры»."""
        outline = self.b.outline()
        kinds = {it.kind for it in outline}
        has_sub = Kind.SUBCHAPTER in kinds
        out = []
        for kind, label in ((Kind.TITLE, "Титульный лист"), (Kind.TOC, "Оглавление"),
                            (Kind.INTRODUCTION, "Введение"), (Kind.CHAPTER, "Глава"),
                            (Kind.SUBCHAPTER, "Подглава"), (Kind.SECTION, "Пункт"),
                            (Kind.CONCLUSION, "Заключение"), (Kind.BIBLIOGRAPHY, "Список литературы"),
                            (Kind.APPENDIX, "Приложение")):
            ok, why = True, ""
            if kind in (Kind.TITLE, Kind.TOC, Kind.INTRODUCTION, Kind.CONCLUSION, Kind.BIBLIOGRAPHY) \
                    and kind in kinds:
                ok, why = False, "уже есть"
            elif kind == Kind.SUBCHAPTER and Kind.CHAPTER not in kinds:
                ok, why = False, "сначала добавьте главу"
            elif kind == Kind.SECTION and not has_sub:
                ok, why = False, "сначала добавьте подглаву"
            out.append((kind, label, ok, why))
        return out

    def _enclosing(self, element, kind: Kind):
        if element is None:
            return None
        i, outline = self._item(element)
        level = LEVEL_KINDS.index(kind) if kind in LEVEL_KINDS else 0
        for k in range(i, -1, -1):
            it = outline[k]
            if it.kind == kind:
                return it
            if it.level < level or (it.level == 0 and it.kind not in LEVEL_KINDS):
                return None
        return None

    def add_element(self, kind: Kind, near=None) -> OpResult:
        """Создать элемент рядом с выбранным пунктом (OpResult.element — новый заголовок)."""
        from .document_builder import BlockType, InsertPosition, InsertResult

        before = {id(it.element) for it in self.b.outline()}
        if kind == Kind.TITLE:
            from .title_page import TitlePageData

            self.b.set_title_page(TitlePageData())
        elif kind == Kind.TOC:
            self.b.insert_block(BlockType.TOC)
        elif kind in (Kind.INTRODUCTION, Kind.CONCLUSION):
            if self.b.find(kind) is not None:
                return OpResult(False, "Такой элемент уже есть")
            block = BlockType.INTRODUCTION if kind == Kind.INTRODUCTION else BlockType.CONCLUSION
            where, anchor = self.b._resolve_position(block, InsertPosition(), InsertResult())
            title = self.b._structural_text(self.b.profile.introduction_title if block ==
                                            BlockType.INTRODUCTION else self.b.profile.conclusion_title)
            self.b._inserter(where, anchor)(Role.STRUCTURAL, title)
            self.b.renumber()
        elif kind == Kind.BIBLIOGRAPHY:
            self.b.bibliography._ensure_heading()
            self.b.renumber()
        elif kind == Kind.APPENDIX:
            self.b.insert_block(BlockType.APPENDIX, heading="Новое приложение")
        elif kind in LEVEL_KINDS:
            block = {Kind.CHAPTER: BlockType.CHAPTER, Kind.SUBCHAPTER: BlockType.SUBCHAPTER,
                     Kind.SECTION: BlockType.SECTION}[kind]
            title = {Kind.CHAPTER: "Новая глава", Kind.SUBCHAPTER: "Новая подглава",
                     Kind.SECTION: "Новый пункт"}[kind]
            pos = InsertPosition()
            if near is not None:
                if kind == Kind.CHAPTER:
                    anchor_item = self._enclosing(near, Kind.CHAPTER)
                elif kind == Kind.SUBCHAPTER:
                    anchor_item = self._enclosing(near, Kind.SUBCHAPTER) or self._enclosing(near, Kind.CHAPTER)
                else:
                    anchor_item = self._enclosing(near, Kind.SECTION) or self._enclosing(near, Kind.SUBCHAPTER)
                if anchor_item is not None:
                    pos = InsertPosition("section_end", anchor_item.element)
            self.b.insert_block(block, heading=title, position=pos)
        else:
            return OpResult(False, "Этот элемент нельзя добавить")
        new = [it for it in self.b.outline() if id(it.element) not in before and it.kind == kind]
        return OpResult(True, f"Добавлено: {new[0].text if new else KIND_RU.get(kind, '')}",
                        element=new[0].element if new else None)

    # ========================================================= удаление
    def describe(self, element) -> str:
        i, outline = self._item(element)
        item = outline[i]
        els = self.section_elements(element)
        subs = sum(1 for it in outline[i + 1:] if it.level > item.level
                   and self._top(it.element) in els)
        paras = sum(1 for el in els[1:] if el.tag == ooxml.W_P and ooxml.element_text(el).strip())
        tables = sum(1 for el in els if el.tag == ooxml.W_TBL)
        parts = [f"абзацев: {paras}"]
        if subs:
            parts.append(f"подпунктов: {subs}")
        if tables and item.kind != Kind.TITLE:
            parts.append(f"таблиц: {tables}")
        return f"«{item.text}» ({KIND_RU.get(item.kind, 'элемент')}; " + ", ".join(parts) + ")"

    def delete(self, element) -> OpResult:
        i, outline = self._item(element)
        item = outline[i]
        if item.element.tag == ooxml.W_P and self.is_locked(item.element):
            return OpResult(False, f"«{item.text}» закреплён — снимите замок, чтобы удалить")
        refs_before = sum(self.b.bibliography.citation_counts().values()) \
            if item.kind == Kind.BIBLIOGRAPHY else 0
        for el in self.section_elements(element):
            el.getparent().remove(el)
        self.b.renumber()
        warnings = []
        if refs_before:
            warnings.append(f"Ссылок на литературу в тексте: {refs_before} — они показывают «[?]»")
        return OpResult(True, f"Удалено: {item.text}", warnings)


__all__ = ["StructureOps", "SectionContent", "OpResult", "has_refs", "make_run"]
