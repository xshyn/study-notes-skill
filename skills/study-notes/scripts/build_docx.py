#!/usr/bin/env python3
"""
build_docx.py — render a study-notes JSON file into a polished, detailed Word document.

  python build_docx.py notes.json -o "Topic - Study Notes.docx"

Layout: cover → at-a-glance (overview, objectives, prerequisites) → clickable contents →
numbered sections with callouts, worked examples, formulas, code, tables, figures and
diagrams → key takeaways → common mistakes → exam tips → cheat sheet → glossary →
review questions → answer key → flashcards → further reading.
Supports right-to-left languages (fa, ar, he, ur …) automatically via meta.language.
"""
from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor, Emu

sys.path.insert(0, str(Path(__file__).resolve().parent))
from notes_common import (CALLOUTS, clean_option, render_chart, strip_step_number, has_rtl, priority_rows, image_size, inline_runs, item_children, item_text,  # noqa: E402
                          labels, load_notes, plain, render_latex, render_mermaid, theme,
                          walk_sections)

CONTENT_W_IN = 6.5


class Builder:
    def __init__(self, notes: dict, base_dir: Path):
        self.n = notes
        self.meta = notes["meta"]
        self.L = labels(self.meta)
        self.T = theme(self.meta)
        self.rtl = self.meta.get("direction") == "rtl"
        self.base_dir = base_dir
        self.tmp = Path(tempfile.mkdtemp(prefix="studynotes_"))
        self.doc = Document()
        self.font = self.meta.get("font_body") or ("Vazirmatn" if self.rtl else "Calibri")
        self.hfont = self.meta.get("font_heading") or self.font
        self.cs_font = self.meta.get("font_complex") or ("Vazirmatn" if self.rtl else self.font)
        self.code_font = "Consolas"
        self._bm = 0
        self._img = 0
        self._setup()

    # ------------------------------------------------------------------ setup / styles
    def _rgb(self, hexstr: str) -> RGBColor:
        return RGBColor.from_string(hexstr.upper())

    def _set_fonts(self, style_or_run, name, cs=None):
        style_or_run.font.name = name
        rpr = style_or_run.element.get_or_add_rPr() if hasattr(style_or_run.element, "get_or_add_rPr") \
            else style_or_run.element.rPr
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        for a in ("w:ascii", "w:hAnsi", "w:eastAsia"):
            rfonts.set(qn(a), name)
        rfonts.set(qn("w:cs"), cs or self.cs_font)

    def _setup(self):
        d = self.doc
        sec = d.sections[0]
        sec.page_width, sec.page_height = Inches(8.5), Inches(11)
        if (self.meta.get("paper") or "").lower() == "a4":
            sec.page_width, sec.page_height = Emu(7560310), Emu(10692130)
        for side in ("left_margin", "right_margin"):
            setattr(sec, side, Inches(1))
        sec.top_margin, sec.bottom_margin = Inches(0.9), Inches(0.9)
        global CONTENT_W_IN
        CONTENT_W_IN = (sec.page_width - sec.left_margin - sec.right_margin) / 914400

        st = d.styles
        normal = st["Normal"]
        self._set_fonts(normal, self.font)
        normal.font.size = Pt(11)
        normal.font.color.rgb = self._rgb(self.T["ink"])
        pf = normal.paragraph_format
        pf.space_after = Pt(6)
        pf.line_spacing = 1.18
        if self.rtl:
            ppr = normal.element.get_or_add_pPr()
            ppr.append(OxmlElement("w:bidi"))
            rpr = normal.element.get_or_add_rPr()
            rpr.append(OxmlElement("w:rtl"))

        sizes = {1: 20, 2: 15, 3: 12.5, 4: 11.5}
        for lvl, size in sizes.items():
            h = st[f"Heading {lvl}"]
            self._set_fonts(h, self.hfont)
            h.font.size = Pt(size)
            h.font.bold = True
            h.font.italic = False
            h.font.color.rgb = self._rgb(self.T["primary"] if lvl < 3 else self.T["ink"])
            h.paragraph_format.space_before = Pt(18 if lvl == 1 else 12)
            h.paragraph_format.space_after = Pt(6)
            h.paragraph_format.keep_with_next = True

        for name in ("List Bullet", "List Bullet 2", "List Bullet 3", "List Number", "List Number 2"):
            try:
                s = st[name]
                s.paragraph_format.space_after = Pt(3)
            except KeyError:
                pass

        self._header_footer(sec)

    def _header_footer(self, sec):
        sec.different_first_page_header_footer = True
        hp = sec.header.paragraphs[0]
        hp.text = ""
        r = hp.add_run(plain(self.meta.get("title", "")))
        r.font.size = Pt(8.5)
        r.font.color.rgb = self._rgb(self.T["muted"])
        hp.alignment = WD_ALIGN_PARAGRAPH.LEFT if not self.rtl else None  # RTL: start = right
        self._border(hp, "bottom", self.T["soft"], 6)
        fp = sec.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = fp.add_run()
        r.font.size = Pt(9)
        r.font.color.rgb = self._rgb(self.T["muted"])
        for tag, txt in (("begin", None), (None, "PAGE"), ("end", None)):
            if tag:
                el = OxmlElement("w:fldChar")
                el.set(qn("w:fldCharType"), tag)
            else:
                el = OxmlElement("w:instrText")
                el.set(qn("xml:space"), "preserve")
                el.text = txt
            r._r.append(el)

    # ------------------------------------------------------------------ low-level helpers
    def _border(self, par, side, color, size=8, space=4):
        ppr = par._p.get_or_add_pPr()
        pbdr = ppr.find(qn("w:pBdr"))
        if pbdr is None:
            pbdr = OxmlElement("w:pBdr")
            ppr.append(pbdr)
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), str(size))
        b.set(qn("w:space"), str(space))
        b.set(qn("w:color"), color)
        pbdr.append(b)

    def _shade_cell(self, cell, fill):
        tcpr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
        tcpr.append(shd)

    def _cell_borders(self, cell, color="FFFFFF", sides=("top", "left", "bottom", "right"), size=4,
                      left_bar=None):
        tcpr = cell._tc.get_or_add_tcPr()
        b = OxmlElement("w:tcBorders")
        for s in ("top", "left", "bottom", "right"):
            el = OxmlElement(f"w:{s}")
            bar_side = ("right" if self.rtl else "left")
            if left_bar and s == bar_side:
                el.set(qn("w:val"), "single")
                el.set(qn("w:sz"), "24")
                el.set(qn("w:color"), left_bar)
            elif s in sides:
                el.set(qn("w:val"), "single")
                el.set(qn("w:sz"), str(size))
                el.set(qn("w:color"), color)
            else:
                el.set(qn("w:val"), "nil")
            b.append(el)
        tcpr.append(b)

    def _cell_margins(self, cell, top=90, bottom=90, start=140, end=140):
        tcpr = cell._tc.get_or_add_tcPr()
        mar = OxmlElement("w:tcMar")
        for k, v in (("top", top), ("bottom", bottom), ("start", start), ("end", end),
                     ("left", start), ("right", end)):
            el = OxmlElement(f"w:{k}")
            el.set(qn("w:w"), str(v))
            el.set(qn("w:type"), "dxa")
            mar.append(el)
        tcpr.append(mar)

    def _table_rtl(self, table):
        if self.rtl:
            tblpr = table._tbl.tblPr
            tblpr.append(OxmlElement("w:bidiVisual"))

    def _par_dir(self, p):
        if self.rtl:
            ppr = p._p.get_or_add_pPr()
            bidi = ppr.find(qn("w:bidi"))
            if bidi is None:
                bidi = OxmlElement("w:bidi")
                ppr.append(bidi)
            if p.text.strip() and not has_rtl(p.text):
                # an all-Latin line (code identifiers, an English answer) reads correctly only as LTR;
                # keep it right-aligned so the page still flows right-to-left
                bidi.set(qn("w:val"), "0")
                if p.alignment in (None, WD_ALIGN_PARAGRAPH.LEFT):
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                bidi.attrib.pop(qn("w:val"), None)
                # in a right-to-left paragraph "left"/"right" are logical (start/end) in Word and
                # LibreOffice; the default (start) is already the right-hand side
                if p.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                    p.alignment = None

    def runs(self, p, text, size=None, color=None, bold=None, italic=None):
        for t, fmt in inline_runs(text):
            r = p.add_run(t)
            if fmt.get("code"):
                self._set_fonts(r, self.code_font, self.code_font)
                r.font.size = Pt((size or 11) - 1)
                on_dark = (color or "").upper() == "FFFFFF"
                r.font.color.rgb = self._rgb("FFD58A" if on_dark else "A61B4A")
                if not on_dark:
                    rpr = r._r.get_or_add_rPr()
                    shd = OxmlElement("w:shd")
                    shd.set(qn("w:val"), "clear")
                    shd.set(qn("w:fill"), "F3F3F6")
                    rpr.append(shd)
            else:
                if size:
                    r.font.size = Pt(size)
                if color:
                    r.font.color.rgb = self._rgb(color)
            if fmt.get("bold") or bold:
                r.bold = True
            if fmt.get("italic") or italic:
                r.italic = True
            if fmt.get("highlight"):
                rpr = r._r.get_or_add_rPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:fill"), "FFF0B3")
                rpr.append(shd)
            if self.rtl and not fmt.get("code") and has_rtl(t):
                r._r.get_or_add_rPr().append(OxmlElement("w:rtl"))
        return p

    def para(self, text="", style=None, container=None, **kw):
        c = container if container is not None else self.doc
        p = c.add_paragraph(style=style) if style else c.add_paragraph()
        self.runs(p, text, **kw)
        self._par_dir(p)
        return p

    def bookmark(self, p, name):
        self._bm += 1
        s = OxmlElement("w:bookmarkStart")
        s.set(qn("w:id"), str(self._bm))
        s.set(qn("w:name"), name)
        e = OxmlElement("w:bookmarkEnd")
        e.set(qn("w:id"), str(self._bm))
        p._p.insert(0, s)
        p._p.append(e)

    def internal_link(self, p, text, anchor, size=11, color=None, bold=False):
        h = OxmlElement("w:hyperlink")
        h.set(qn("w:anchor"), anchor)
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        if color:
            c = OxmlElement("w:color")
            c.set(qn("w:val"), color)
            rpr.append(c)
        if bold:
            rpr.append(OxmlElement("w:b"))
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(size * 2)))
        rpr.append(sz)
        if self.rtl:
            rpr.append(OxmlElement("w:rtl"))
        r.append(rpr)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        r.append(t)
        h.append(r)
        p._p.append(h)

    def heading(self, text, level, anchor=None):
        p = self.doc.add_heading(level=level)
        if getattr(self, "_pb", False):
            p.paragraph_format.page_break_before = True
            self._pb = False
        self.runs(p, text)
        self._par_dir(p)
        if level == 1:
            self._border(p, "bottom", self.T["accent"], 12, 4)
        if anchor:
            self.bookmark(p, anchor)
        return p

    def page_break(self):
        # applied to the next heading as "page break before" — never leaves a blank page behind
        self._pb = True

    def box(self, fill, bar=None, width_in=None, border_color=None):
        """A one-cell table used as a callout container. Returns the cell."""
        t = self.doc.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        self._table_rtl(t)
        cell = t.cell(0, 0)
        cell.width = Inches(width_in or CONTENT_W_IN)
        # keep callouts in one piece: never leave a lone label at the bottom of a page
        t.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        self._shade_cell(cell, fill)
        self._cell_borders(cell, color=border_color or fill, left_bar=bar)
        self._cell_margins(cell)
        cell.paragraphs[0].paragraph_format.space_after = Pt(2)
        return cell

    def spacer(self, pts=4):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        r = p.add_run()
        r.font.size = Pt(pts)

    def cell_text(self, cell, text, first=True, **kw):
        p = cell.paragraphs[0] if first and not cell.paragraphs[0].text else cell.add_paragraph()
        self.runs(p, text, **kw)
        self._par_dir(p)
        p.paragraph_format.space_after = Pt(3)
        return p

    def picture(self, path: Path, width_in: float, caption: str | None = None, container=None):
        w, h = image_size(path)
        max_h = 7.2
        if width_in * h / max(w, 1) > max_h:
            width_in = max_h * w / h
        c = container if container is not None else self.doc
        p = c.add_paragraph() if container is None else (c.paragraphs[0] if not c.paragraphs[0].text else c.add_paragraph())
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(path), width=Inches(width_in))
        if caption:
            cp = c.add_paragraph()
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self.runs(cp, caption, size=9, color=self.T["muted"], italic=True)

    # ------------------------------------------------------------------ blocks
    def bullets(self, items, numbered=False, level=0, container=None):
        if numbered:
            for i, it in enumerate(items, 1):
                c = container if container is not None else self.doc
                p = c.add_paragraph()
                pf = p.paragraph_format
                pf.space_after = Pt(3)
                ind = Inches(0.3 + 0.3 * level)
                if self.rtl:
                    pf.right_indent = ind
                else:
                    pf.left_indent = ind
                pf.first_line_indent = Inches(-0.25)
                self.runs(p, f"**{i}.** ")
                self.runs(p, item_text(it))
                self._par_dir(p)
                kids = item_children(it)
                if kids:
                    self.bullets(kids, numbered=False, level=level + 1, container=container)
            return
        base = "List Number" if numbered else "List Bullet"
        style = base if level == 0 else f"{base} {min(level + 1, 3 if not numbered else 2)}"
        for it in items:
            if container is not None:
                p = container.add_paragraph(style=style)
            else:
                p = self.doc.add_paragraph(style=style)
            self.runs(p, item_text(it))
            self._par_dir(p)
            kids = item_children(it)
            if kids:
                self.bullets(kids, numbered=False, level=level + 1, container=container)

    def block(self, b: dict):
        t = (b.get("type") or "text").lower()
        fn = getattr(self, f"b_{t}", None)
        if fn is None:
            alias = {"paragraph": "text", "list": "bullets", "steps": "numbered", "numbered_list": "numbered",
                     "callout": "key_point", "important": "key_point", "pitfall": "warning",
                     "mistake": "warning", "exam_tip": "tip", "diagram_mermaid": "diagram",
                     "figure": "image", "math": "formula", "equation": "formula", "graph": "chart", "plot": "chart"}.get(t)
            fn = getattr(self, f"b_{alias}", None) if alias else None
        if fn is None:
            fn = self.b_text
        fn(b)

    def b_text(self, b):
        self.para(b.get("text", ""))

    def b_bullets(self, b):
        if b.get("title"):
            self.para(b["title"], bold=True)
        self.bullets(b.get("items", []))

    def b_numbered(self, b):
        if b.get("title"):
            self.para(b["title"], bold=True)
        self.bullets(b.get("items", []), numbered=True)

    def _callout(self, kind, text, title=None, items=None):
        bg, bar, _ = CALLOUTS.get(kind, CALLOUTS["note"])
        label = title or self.L.get(kind, kind)
        cell = self.box(bg, bar=bar)
        self.cell_text(cell, label.upper() if self.meta["language"] == "en" else label, size=8.5,
                       color=bar, bold=True)
        if text:
            self.cell_text(cell, text, first=False)
        if items:
            for it in items:
                p = cell.add_paragraph(style="List Bullet")
                self.runs(p, item_text(it))
                self._par_dir(p)
        self.spacer()

    def b_key_point(self, b):
        self._callout("key_point", b.get("text", ""), b.get("title"), b.get("items"))

    def b_warning(self, b):
        self._callout("warning", b.get("text", ""), b.get("title"), b.get("items"))

    def b_tip(self, b):
        self._callout("tip", b.get("text", ""), b.get("title"), b.get("items"))

    def b_mnemonic(self, b):
        self._callout("mnemonic", b.get("text", ""), b.get("title"), b.get("items"))

    def b_note(self, b):
        self._callout("note", b.get("text", ""), b.get("title"), b.get("items"))

    def b_intuition(self, b):
        self._callout("intuition", b.get("text", ""), b.get("title"), b.get("items"))

    def b_definition(self, b):
        cell = self.box(self.T["soft"], bar=self.T["primary"])
        p = self.cell_text(cell, "", size=11)
        self.runs(p, b.get("term", ""), bold=True, color=self.T["primary"], size=11.5)
        if b.get("aka"):
            self.runs(p, f"  ({b['aka']})", color=self.T["muted"], size=9.5, italic=True)
        self.cell_text(cell, b.get("definition", b.get("text", "")), first=False)
        if b.get("note"):
            self.cell_text(cell, b["note"], first=False, size=9.5, color=self.T["muted"], italic=True)
        if b.get("example"):
            self.cell_text(cell, f"*e.g.* {b['example']}", first=False, size=10)
        self.spacer()

    def b_example(self, b):
        cell = self.box("FAFAFC", bar=self.T["accent"], border_color="E3E4EA")
        title = b.get("title") or self.L["example"]
        self.cell_text(cell, f"{self.L['example']}: {title}" if b.get("title") else title,
                       bold=True, color=self.T["primary"])
        if b.get("problem"):
            p = self.cell_text(cell, "", first=False)
            self.runs(p, f"{self.L['problem']}. ", bold=True)
            self.runs(p, b["problem"])
        if b.get("text"):
            self.cell_text(cell, b["text"], first=False)
        steps = b.get("steps") or []
        if steps:
            if b.get("problem"):
                self.cell_text(cell, f"**{self.L['solution']}.**", first=False)
            for i, s in enumerate(steps, 1):
                p = cell.add_paragraph()
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.left_indent = Inches(0.15)
                self.runs(p, f"**{i}.** ")
                self.runs(p, strip_step_number(item_text(s)))
                self._par_dir(p)
        if b.get("code"):
            self._code_into(cell, b["code"])
        if b.get("solution"):
            p = self.cell_text(cell, "", first=False)
            if not steps:
                self.runs(p, f"{self.L['solution']}. ", bold=True)
            self.runs(p, b["solution"])
        if b.get("answer"):
            p = self.cell_text(cell, "", first=False)
            self.runs(p, "∴ ", bold=True, color=self.T["accent"])
            self.runs(p, b["answer"], bold=True)
        self.spacer()

    def b_formula(self, b):
        cell = self.box("FFFFFF", border_color="D9DBE5")
        if b.get("name"):
            self.cell_text(cell, b["name"], bold=True, size=9.5, color=self.T["muted"])
        shown = False
        if b.get("latex"):
            self._img += 1
            png = self.tmp / f"f{self._img}.png"
            if render_latex(b["latex"], png, color=self.T["ink"]):
                w, h = image_size(png)
                width = min(CONTENT_W_IN - 0.6, w / 220 * 0.9)
                p = cell.paragraphs[0] if not cell.paragraphs[0].text and not b.get("name") else cell.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(str(png), width=Inches(width))
                shown = True
        if not shown:
            p = cell.paragraphs[0] if not cell.paragraphs[0].text and not b.get("name") else cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(b.get("text") or b.get("formula") or b.get("latex", ""))
            self._set_fonts(r, "Cambria Math", "Cambria Math")
            r.font.size = Pt(14)
        if b.get("variables"):
            self.cell_text(cell, f"*{self.L['where']}*", first=False, size=10)
            for v in b["variables"]:
                p = cell.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.25)
                p.paragraph_format.space_after = Pt(1)
                r = p.add_run(v.get("symbol", ""))
                r.bold = True
                r.font.size = Pt(10)
                self.runs(p, f" — {v.get('meaning', '')}", size=10)
                self._par_dir(p)
        if b.get("explanation"):
            self.cell_text(cell, b["explanation"], first=False, size=10.5)
        self.spacer()

    def _code_into(self, cell_or_doc, code, language=None):
        is_doc = cell_or_doc is self.doc
        if is_doc:
            cell = self.box("F6F8FA", border_color="E1E4E8")
        else:
            inner = cell_or_doc.add_table(rows=1, cols=1)
            cell = inner.cell(0, 0)
            self._shade_cell(cell, "F6F8FA")
            self._cell_borders(cell, color="E1E4E8")
            self._cell_margins(cell, 60, 60, 120, 120)
        lines = code.rstrip("\n").split("\n")
        for i, ln in enumerate(lines):
            p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            ppr = p._p.get_or_add_pPr()
            bidi = OxmlElement("w:bidi")
            bidi.set(qn("w:val"), "0")
            ppr.append(bidi)
            r = p.add_run(ln if ln else " ")
            self._set_fonts(r, self.code_font, self.code_font)
            r.font.size = Pt(9)
            r.font.color.rgb = self._rgb("24292E")

    def b_code(self, b):
        if b.get("caption") or b.get("title"):
            self.para(b.get("caption") or b.get("title"), size=9.5, color=self.T["muted"], italic=True)
        self._code_into(self.doc, b.get("code", ""), b.get("language"))
        if b.get("explanation"):
            self.para(b["explanation"])
        else:
            self.spacer()

    def b_table(self, b):
        headers = b.get("headers") or []
        rows = b.get("rows") or []
        if b.get("caption") or b.get("title"):
            self.para(b.get("caption") or b.get("title"), bold=True, size=10,
                      color=self.T["primary"]).paragraph_format.keep_with_next = True
        ncols = max([len(headers)] + [len(r) for r in rows] + [1])
        t = self.doc.add_table(rows=0, cols=ncols)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        self._table_rtl(t)
        t.autofit = True
        if headers:
            cells = t.add_row().cells
            for i, h in enumerate(headers):
                self._shade_cell(cells[i], self.T["primary"])
                self._cell_borders(cells[i], color=self.T["primary"])
                self._cell_margins(cells[i], 60, 60, 100, 100)
                self.cell_text(cells[i], str(h), bold=True, color="FFFFFF", size=10)
            tr_pr = t.rows[0]._tr.get_or_add_trPr()
            th = OxmlElement("w:tblHeader")
            tr_pr.append(th)
        for ri, row in enumerate(rows):
            new_row = t.add_row()
            new_row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
            cells = new_row.cells
            for i in range(ncols):
                v = row[i] if i < len(row) else ""
                fill = self.T["soft"] if ri % 2 else "FFFFFF"
                self._shade_cell(cells[i], fill)
                self._cell_borders(cells[i], color="DADCE6")
                self._cell_margins(cells[i], 50, 50, 100, 100)
                self.cell_text(cells[i], str(v), size=10, bold=(i == 0 and b.get("first_col_bold", True)
                                                               and ncols > 2))
        if b.get("note"):
            self.para(b["note"], size=9, color=self.T["muted"], italic=True)
        else:
            self.spacer(6)

    def b_quote(self, b):
        p = self.para(f"“{b.get('text', '')}”", italic=True, color=self.T["muted"])
        p.paragraph_format.left_indent = Inches(0.4)
        self._border(p, "right" if self.rtl else "left", self.T["accent"], 18, 8)
        if b.get("by"):
            self.para(f"— {b['by']}", size=9.5, color=self.T["muted"]).paragraph_format.left_indent = Inches(0.4)

    def _resolve(self, path_str):
        p = Path(path_str).expanduser()
        if not p.is_absolute():
            p = (self.base_dir / p).resolve()
        return p

    def b_image(self, b):
        p = self._resolve(b.get("path", ""))
        if p.exists():
            self.picture(p, min(b.get("width_in", 5.5), CONTENT_W_IN), b.get("caption"))
        elif b.get("caption"):
            self.para(f"[Figure: {b['caption']}]", italic=True, color=self.T["muted"])

    def b_chart(self, b):
        self._img += 1
        png = self.tmp / f"c{self._img}.png"
        if render_chart(b, png, self.T, rtl=self.rtl):
            self.picture(png, min(b.get("width_in", 6.0), CONTENT_W_IN), b.get("caption"))
        elif b.get("caption"):
            self.para(f"[Chart: {b['caption']}]", italic=True, color=self.T["muted"])

    def b_diagram(self, b):
        self._img += 1
        png = self.tmp / f"d{self._img}.png"
        code = b.get("mermaid") or b.get("code") or ""
        if code and render_mermaid(code, png):
            self.picture(png, min(b.get("width_in", 5.8), CONTENT_W_IN), b.get("caption"))
        elif b.get("ascii"):
            self._code_into(self.doc, b["ascii"])
            if b.get("caption"):
                self.para(b["caption"], size=9, italic=True, color=self.T["muted"])
        else:
            self.b_code({"code": code, "caption": b.get("caption")})

    # ------------------------------------------------------------------ document parts
    def cover(self):
        m = self.meta
        for _ in range(5):
            self.spacer(14)
        p = self.para(m.get("kicker") or self.L["kicker"], size=10, color=self.T["accent"], bold=True)
        p.paragraph_format.space_after = Pt(4)
        p = self.para(m.get("title", "Study Notes"), size=32, color=self.T["primary"], bold=True)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_after = Pt(6)
        if m.get("subtitle"):
            self.para(m["subtitle"], size=15, color=self.T["muted"])
        bar = self.para("")
        self._border(bar, "bottom", self.T["accent"], 24, 1)
        bar.paragraph_format.space_after = Pt(18)
        rows = []
        if m.get("source"):
            rows.append((self.L["source"], m["source"]))
        for k, lab in (("author", "Author"), ("course", "Course"), ("level", "Level"), ("date", "Date")):
            if m.get(k):
                rows.append((self.L.get(k, lab), str(m[k])))
        for k, v in rows:
            p = self.para("", size=10.5)
            self.runs(p, f"{k}:  ", bold=True, color=self.T["ink"], size=10.5)
            self.runs(p, v, color=self.T["muted"], size=10.5)
            p.paragraph_format.space_after = Pt(2)
        stats = self._stats()
        self.spacer(18)
        p = self.para(stats, size=9.5, color=self.T["muted"])
        self.page_break()

    def _stats(self):
        n = self.n
        nsec = len(n.get("sections", []))
        parts = [f"{nsec} sections"]
        for key, lab in (("glossary", "glossary terms"), ("review_questions", "practice questions"),
                         ("flashcards", "flashcards")):
            if n.get(key):
                parts.append(f"{len(n[key])} {lab}")
        return "  ·  ".join(parts) if self.meta["language"] == "en" else ""

    def at_a_glance(self):
        n = self.n
        if n.get("overview"):
            self.heading(self.L["overview"], 1, "sec_overview")
            for para in (n["overview"] if isinstance(n["overview"], list) else [n["overview"]]):
                self.para(para)
        if n.get("exam_priorities"):
            self.heading(self.L["priorities"], 2)
            self.b_table({"headers": [self.L["topic"], self.L["priority"], self.L["focus"]],
                          "rows": priority_rows(n, self.L), "first_col_bold": True})
        if n.get("learning_objectives"):
            self.heading(self.L["objectives"], 2)
            self.para("After studying these notes you should be able to:" if self.meta["language"] == "en" else "",
                      italic=True, color=self.T["muted"]) if self.meta["language"] == "en" else None
            self.bullets(n["learning_objectives"])
        if n.get("prerequisites"):
            self.heading(self.L["prereq"], 2)
            self.bullets(n["prerequisites"])

    def contents(self):
        self.heading(self.L["contents"], 1)
        entries = []
        for num, depth, s in walk_sections(self.n["sections"]):
            if depth <= 2:
                entries.append((depth, f"{num}  {plain(s.get('title', ''))}", f"sec_{num.replace('.', '_')}"))
        for key, lab in (("key_takeaways", "takeaways"), ("common_mistakes", "mistakes"),
                         ("exam_tips", "exam_tips"), ("cheat_sheet", "cheat"), ("glossary", "glossary"),
                         ("review_questions", "questions"), ("flashcards", "flashcards")):
            if self.n.get(key):
                entries.append((1, self.L[lab], f"end_{key}"))
        for depth, text, anchor in entries:
            p = self.doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2 if depth > 1 else 4)
            p.paragraph_format.left_indent = Inches(0.3 * (depth - 1))
            self._par_dir(p)
            self.internal_link(p, text, anchor, size=10.5 if depth > 1 else 11.5,
                               color=self.T["ink"] if depth > 1 else self.T["primary"], bold=depth == 1)
        self.page_break()

    def sections(self):
        for num, depth, s in walk_sections(self.n["sections"]):
            lvl = min(depth, 3)
            if depth == 1 and num != "1" and self.meta.get("section_page_breaks", True):
                self.page_break()
            self.heading(f"{num}  {s.get('title', '')}", lvl, f"sec_{num.replace('.', '_')}")
            if s.get("source_ref"):
                self.para(f"{self.L['source']}: {s['source_ref']}", size=8.5, color=self.T["muted"], italic=True)
            if s.get("summary"):
                cell = self.box(self.T["soft"])
                self.cell_text(cell, self.L["in_brief"].upper() if self.meta["language"] == "en"
                               else self.L["in_brief"], size=8.5, bold=True, color=self.T["primary"])
                self.cell_text(cell, s["summary"], first=False, italic=True)
                self.spacer(6)
            for b in s.get("blocks", []):
                self.block(b)
            if s.get("takeaways"):
                self._callout("key_point", "", self.L["takeaways"], s["takeaways"])

    def back_matter(self):
        n, L = self.n, self.L
        if n.get("key_takeaways"):
            self.page_break()
            self.heading(L["takeaways"], 1, "end_key_takeaways")
            self.bullets(n["key_takeaways"], numbered=True)
        if n.get("common_mistakes"):
            self.heading(L["mistakes"], 1, "end_common_mistakes")
            for m in n["common_mistakes"]:
                if isinstance(m, dict):
                    cell = self.box(CALLOUTS["warning"][0], bar=CALLOUTS["warning"][1])
                    p = self.cell_text(cell, "")
                    self.runs(p, "✗ ", bold=True, color="C0392B")
                    self.runs(p, m.get("mistake", m.get("text", "")))
                    if m.get("correction") or m.get("fix"):
                        p2 = self.cell_text(cell, "", first=False)
                        self.runs(p2, "✓ ", bold=True, color="27AE60")
                        self.runs(p2, m.get("correction") or m.get("fix"))
                    self.spacer(3)
                else:
                    self.bullets([m])
        if n.get("exam_tips"):
            self.heading(L["exam_tips"], 1, "end_exam_tips")
            self._callout("tip", "", L["exam_tips"], n["exam_tips"])
        if n.get("cheat_sheet"):
            self.page_break()
            self.heading(L["cheat"], 1, "end_cheat_sheet")
            self.cheat_sheet(n["cheat_sheet"])
        if n.get("glossary"):
            self.page_break()
            self.heading(L["glossary"], 1, "end_glossary")
            items = sorted(n["glossary"], key=lambda g: plain(g.get("term", "")).lower())
            self.b_table({"headers": [L["term"], L["definition"]],
                          "rows": [[f"**{g.get('term', '')}**", g.get("definition", "")] for g in items],
                          "first_col_bold": False})
        if n.get("review_questions"):
            self.page_break()
            self.questions(n["review_questions"])
        if n.get("flashcards"):
            self.page_break()
            self.heading(L["flashcards"], 1, "end_flashcards")
            self.para("Cover the right column and test yourself." if self.meta["language"] == "en" else "",
                      italic=True, color=self.T["muted"], size=9.5)
            self.b_table({"headers": [L["front"], L["back"]],
                          "rows": [[f.get("front", ""), f.get("back", "")] for f in n["flashcards"]],
                          "first_col_bold": False})
        if n.get("further_reading"):
            self.heading(L["further"], 1)
            self.bullets(n["further_reading"])

    def cheat_sheet(self, groups):
        # two-column grid of compact boxes
        groups = [g if isinstance(g, dict) else {"heading": "", "items": [g]} for g in groups]
        t = self.doc.add_table(rows=0, cols=2)
        self._table_rtl(t)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i in range(0, len(groups), 2):
            cells = t.add_row().cells
            for j in range(2):
                c = cells[j]
                c.width = Inches(CONTENT_W_IN / 2)
                self._cell_margins(c, 80, 80, 110, 110)
                self._cell_borders(c, color="FFFFFF", size=12)
                if i + j >= len(groups):
                    continue
                g = groups[i + j]
                self._shade_cell(c, self.T["soft"])
                self.cell_text(c, g.get("heading", ""), bold=True, color=self.T["primary"], size=10.5)
                for it in g.get("items", []):
                    p = c.add_paragraph()
                    p.paragraph_format.space_after = Pt(1.5)
                    self.runs(p, "• ", color=self.T["accent"], bold=True, size=9.5)
                    self.runs(p, item_text(it), size=9.5)
                    self._par_dir(p)

    @staticmethod
    def _split_code(text):
        """Multi-line Q/A text: first line stays prose; the rest is code (LTR box).
        A ```fenced``` block is always treated as code."""
        text = text or ""
        if "```" in text:
            pre, _, rest = text.partition("```")
            code = rest.split("\n", 1)[1] if "\n" in rest else rest
            code = code.split("```")[0]
            return pre.strip(), code.rstrip()
        if "\n" not in text:
            return text, None
        head, rest = text.split("\n", 1)
        if head.lstrip().startswith(("class ", "def ", "import ", "from ")):
            return "", text
        return head, rest

    def questions(self, qs):
        L = self.L
        self.heading(L["questions"], 1, "end_review_questions")
        if self.meta["language"] == "en":
            self.para("Answer from memory first, then check the answer key. Questions are grouped by "
                      "type and increase in difficulty.", italic=True, color=self.T["muted"], size=9.5)
        order = {"recall": 0, "true_false": 1, "mcq": 2, "conceptual": 3, "application": 4, "problem": 5,
                 "compare": 3, "analysis": 6}
        diff = {"easy": 0, "medium": 1, "hard": 2}
        indexed = list(enumerate(qs, 1))
        groups: dict[str, list] = {}
        for i, q in indexed:
            groups.setdefault(q.get("type", "conceptual"), []).append((i, q))
        numbering = {}
        k = 0
        for gtype in sorted(groups, key=lambda g: order.get(g, 9)):
            items = sorted(groups[gtype], key=lambda x: diff.get(x[1].get("difficulty", "medium"), 1))
            self.heading(L.get(f"q_{gtype}", gtype.replace("_", " ").title()), 3)
            for _, q in items:
                k += 1
                numbering[id(q)] = k
                p = self.para("")
                self.runs(p, f"Q{k}. ", bold=True, color=self.T["primary"])
                qtext, qcode = self._split_code(q.get("question", ""))
                self.runs(p, qtext)
                if q.get("difficulty"):
                    dot = {"easy": "●○○", "medium": "●●○", "hard": "●●●"}.get(q["difficulty"], "")
                    self.runs(p, f"  {dot}", size=8, color=self.T["accent"])
                self._par_dir(p)
                if qcode:
                    self._code_into(self.doc, qcode)
                if q.get("code"):
                    self._code_into(self.doc, q["code"])
                if q.get("options"):
                    for oi, opt in enumerate(q["options"]):
                        op = self.para(f"{chr(65 + oi)}.  {clean_option(opt)}")
                        op.paragraph_format.left_indent = Inches(0.35)
                        op.paragraph_format.space_after = Pt(1)
        self.page_break()
        self.heading(L["answers"], 1, "end_answers")
        for gtype in sorted(groups, key=lambda g: order.get(g, 9)):
            items = sorted(groups[gtype], key=lambda x: diff.get(x[1].get("difficulty", "medium"), 1))
            for _, q in items:
                k = numbering[id(q)]
                p = self.para("")
                self.runs(p, f"Q{k}. ", bold=True, color=self.T["primary"])
                atext, acode = self._split_code(q.get("answer", ""))
                self.runs(p, atext)
                self._par_dir(p)
                if acode:
                    self._code_into(self.doc, acode)
                if q.get("answer_code"):
                    self._code_into(self.doc, q["answer_code"])
                if q.get("explanation"):
                    self.para(q["explanation"], size=9.5, color=self.T["muted"]).paragraph_format.left_indent = \
                        Inches(0.3)
                if q.get("section"):
                    self.para(f"{L['see']}{q['section']}", size=8.5, color=self.T["muted"], italic=True)

    def build(self, out: Path):
        self.cover()
        self.at_a_glance()
        self.page_break() if (self.n.get("overview") or self.n.get("learning_objectives")) else None
        self.contents()
        self.sections()
        self.back_matter()
        cp = self.doc.core_properties
        cp.title = plain(self.meta.get("title", ""))
        cp.subject = plain(self.meta.get("subtitle", ""))
        cp.author = self.meta.get("author_notes", "study-notes skill")
        cp.keywords = "study notes; " + plain(self.meta.get("title", ""))
        out.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(str(out))
        return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("notes", help="notes JSON file (see references/notes_schema.md)")
    ap.add_argument("-o", "--output", required=True, help="output .docx path")
    a = ap.parse_args()
    notes = load_notes(a.notes)
    out = Builder(notes, Path(a.notes).resolve().parent).build(Path(a.output))
    print(f"[build_docx] wrote {out}")


if __name__ == "__main__":
    main()
