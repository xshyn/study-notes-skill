#!/usr/bin/env python3
"""
build_pptx.py — render a study-notes JSON file into a detailed 16:9 study / review deck.

  python build_pptx.py notes.json -o "Topic - Study Deck.pptx"

Deck flow: title → learning objectives → roadmap → for each section: divider, content slides
(auto-paginated bullets, definition cards, callouts, worked examples, formulas, code, tables,
figures, diagrams) and a "check yourself" mini-quiz with answers → key takeaways → common
mistakes → exam tips → cheat sheet → glossary → final quiz → answers.
Every slide carries speaker notes with the fuller explanation, so the deck doubles as a
lecture script. Text is sized to fit its box and overflow is split onto continuation slides.
RTL languages (fa, ar, he, ur …) are laid out right-to-left automatically.
"""
from __future__ import annotations

import argparse
import math
import re
import sys
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from notes_common import (CALLOUTS, image_size, inline_runs, item_children, item_text,  # noqa: E402
                          labels, load_notes, plain, render_latex, render_mermaid, theme,
                          walk_sections)

W, H = 13.333, 7.5
MARGIN = 0.6
TITLE_TOP = 0.42
BODY_TOP = 1.45
BODY_BOTTOM = H - 0.55
MIN_PT = 13


def rgb(h: str) -> RGBColor:
    return RGBColor.from_string(h.lstrip("#").upper())


def sentences(text: str) -> list[str]:
    parts = re.split(r"(?<=[.!?؟۔])\s+(?=[A-Z0-9؀-ۿ\"“(])", text.strip())
    return [p for p in parts if p]


class Deck:
    def __init__(self, notes: dict, base_dir: Path):
        self.n = notes
        self.meta = notes["meta"]
        self.L = labels(self.meta)
        self.T = theme(self.meta)
        self.rtl = self.meta.get("direction") == "rtl"
        self.base_dir = base_dir
        self.tmp = Path(tempfile.mkdtemp(prefix="studydeck_"))
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(W), Inches(H)
        self.blank = self.prs.slide_layouts[6]
        self.font = self.meta.get("font_body") or ("Vazirmatn" if self.rtl else "Calibri")
        self.hfont = self.meta.get("font_heading") or self.font
        self.code_font = "Consolas"
        self.k = 0
        self.char_w = 0.56 if self.rtl else 0.5  # average glyph width in em
        self.qn = 0
        self.root_title = ""

    # ------------------------------------------------------------------ primitives
    def new_slide(self, notes: str | None = None, dark=False):
        s = self.prs.slides.add_slide(self.blank)
        bg = s.background.fill
        bg.solid()
        bg.fore_color.rgb = rgb(self.T["primary"] if dark else "FFFFFF")
        if notes:
            s.notes_slide.notes_text_frame.text = plain(notes)
        return s

    def add_notes(self, slide, text):
        if not text:
            return
        tf = slide.notes_slide.notes_text_frame
        tf.text = (tf.text + "\n\n" if tf.text else "") + plain(text)

    def rect(self, slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
        shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(fill)
        if line:
            shp.line.color.rgb = rgb(line)
            shp.line.width = Pt(1)
        else:
            shp.line.fill.background()
        shp.shadow.inherit = False
        if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            shp.adjustments[0] = radius
        shp.text_frame.text = ""
        return shp

    def _para_dir(self, p, align=None):
        ppr = p._p.get_or_add_pPr()
        if self.rtl:
            ppr.set("rtl", "1")
            p.alignment = align or PP_ALIGN.RIGHT
        elif align is not None:
            p.alignment = align

    def _runs(self, p, text, size, color, bold=False, italic=False, font=None):
        for t, fmt in inline_runs(text):
            r = p.add_run()
            r.text = t
            f = r.font
            f.size = Pt(size)
            f.name = self.code_font if fmt.get("code") else (font or self.font)
            f.bold = bool(bold or fmt.get("bold"))
            f.italic = bool(italic or fmt.get("italic"))
            if fmt.get("code"):
                f.color.rgb = rgb("A61B4A")
            elif fmt.get("highlight"):
                f.color.rgb = rgb(self.T["accent"])
            else:
                f.color.rgb = rgb(color)
            if self.rtl and not fmt.get("code"):
                rpr = r._r.get_or_add_rPr()
                cs = rpr.makeelement(qn("a:cs"), {"typeface": font or self.font})
                rpr.append(cs)

    def textbox(self, slide, x, y, w, h, paras, size, color=None, bold=False, italic=False,
                anchor=MSO_ANCHOR.TOP, align=None, font=None, spacing=1.12, space_after=None):
        """paras: list of str or (str, level, opts) tuples."""
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = None
        tf.vertical_anchor = anchor
        for side in ("margin_left", "margin_right"):
            setattr(tf, side, Inches(0.05))
        tf.margin_top = tf.margin_bottom = Inches(0.03)
        color = color or self.T["ink"]
        first = True
        for item in paras:
            if isinstance(item, tuple):
                text, level, opts = item
            else:
                text, level, opts = item, 0, {}
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.line_spacing = spacing
            sa = opts.get("space_after", space_after if space_after is not None else size * 0.45)
            p.space_after = Pt(sa)
            psize = opts.get("size", size - 2 * level)
            bullet = opts.get("bullet")
            if bullet:
                self._bullet(p, bullet, level, psize, opts.get("bullet_color", self.T["accent"]))
            self._para_dir(p, opts.get("align", align))
            self._runs(p, text, psize, opts.get("color", color), opts.get("bold", bold),
                       opts.get("italic", italic), opts.get("font", font))
        return tb

    def _bullet(self, p, char, level, size, color):
        ppr = p._p.get_or_add_pPr()
        indent = int(Inches(0.32 + 0.0 * level))
        mar = int(Inches(0.32 + 0.38 * level))
        ppr.set("marL", str(mar))
        ppr.set("indent", str(-indent))
        for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buClr", "a:buSzPct"):
            for el in ppr.findall(qn(tag)):
                ppr.remove(el)
        clr = ppr.makeelement(qn("a:buClr"), {})
        srgb = clr.makeelement(qn("a:srgbClr"), {"val": color})
        clr.append(srgb)
        ppr.append(clr)
        if char == "num":
            ppr.append(ppr.makeelement(qn("a:buAutoNum"), {"type": "arabicPeriod"}))
        else:
            ppr.append(ppr.makeelement(qn("a:buChar"), {"char": char if level == 0 else "–"}))

    # ------------------------------------------------------------------ text fitting
    def est_height(self, paras, width, size, spacing=1.12):
        total = 0.0
        for item in paras:
            text, level, opts = item if isinstance(item, tuple) else (item, 0, {})
            ps = opts.get("size", size - 2 * level)
            indent = 0.32 + 0.38 * level if opts.get("bullet") else 0
            cpl = max(8, (width - indent - 0.1) * 72 / (ps * self.char_w))
            lines = 0
            for seg in plain(text).split("\n"):
                lines += max(1, math.ceil(len(seg) / cpl))
            total += lines * ps * spacing * 1.2 / 72 + opts.get("space_after", ps * 0.45) / 72
        return total

    def fit_size(self, paras, width, height, start=22, minimum=MIN_PT):
        size = start
        while size > minimum and self.est_height(paras, width, size) > height:
            size -= 1
        return size, self.est_height(paras, width, size) <= height

    def split_to_fit(self, paras, width, height, start=22):
        """Greedy pagination: returns list of (paras_chunk, size)."""
        pages, cur = [], []
        for p in paras:
            trial = cur + [p]
            if cur and not self.fit_size(trial, width, height, start)[1]:
                pages.append(cur)
                cur = [p]
            else:
                cur = trial
        if cur:
            pages.append(cur)
        out = []
        for pg in pages:
            size, _ = self.fit_size(pg, width, height, start)
            out.append((pg, size))
        return out

    # ------------------------------------------------------------------ chrome
    def chrome(self, slide, title, kicker=None, continued=False):
        """Standard light slide frame: accent bar, title, footer."""
        x = MARGIN
        self.rect(slide, 0, 0, W, 0.09, self.T["primary"])
        self.rect(slide, (W - MARGIN - 0.9) if self.rtl else MARGIN, 0.09, 0.9, 0.05, self.T["accent"])
        if kicker:
            self.textbox(slide, x, 0.24, W - 2 * MARGIN, 0.3, [kicker.upper() if not self.rtl else kicker],
                         10, color=self.T["accent"], bold=True)
        t = plain(title) + (f"  ({self.L['continued']})" if continued else "")
        tsize, _ = self.fit_size([t], W - 2 * MARGIN, 0.8, 30, 20)
        self.textbox(slide, x, TITLE_TOP + 0.08, W - 2 * MARGIN, 0.85, [t], tsize, color=self.T["primary"],
                     bold=True, font=self.hfont, anchor=MSO_ANCHOR.MIDDLE)
        self.k += 1
        foot = plain(self.meta.get("short_title") or self.meta.get("title", ""))
        self.textbox(slide, MARGIN, H - 0.42, W - 2 * MARGIN - 1, 0.3, [foot], 9, color=self.T["muted"],
                     align=PP_ALIGN.RIGHT if self.rtl else PP_ALIGN.LEFT)
        self.textbox(slide, (MARGIN if self.rtl else W - MARGIN - 1), H - 0.42, 1, 0.3, [str(self.k + 1)], 9,
                     color=self.T["muted"], align=PP_ALIGN.LEFT if self.rtl else PP_ALIGN.RIGHT)

    # ------------------------------------------------------------------ slide types
    def title_slide(self):
        m = self.meta
        s = self.new_slide(dark=True, notes=self.n.get("overview") if isinstance(self.n.get("overview"), str)
                           else " ".join(self.n.get("overview") or []))
        self.rect(s, 0, H - 0.18, W, 0.18, self.T["accent"])
        self.rect(s, MARGIN, 2.05, 1.2, 0.08, self.T["accent"])
        self.textbox(s, MARGIN, 1.45, W - 2 * MARGIN, 0.5, [m.get("kicker") or self.L["kicker"]], 14,
                     color=self.T["accent"], bold=True)
        title = plain(m.get("title", "Study Notes"))
        size, _ = self.fit_size([title], W - 2 * MARGIN, 1.9, 48, 30)
        self.textbox(s, MARGIN, 2.3, W - 2 * MARGIN, 2.0, [title], size, color="FFFFFF", bold=True,
                     font=self.hfont, anchor=MSO_ANCHOR.TOP, spacing=1.0)
        if m.get("subtitle"):
            self.textbox(s, MARGIN, 4.35, W - 2 * MARGIN, 0.9, [m["subtitle"]], 20, color="DDE1F2")
        meta_line = "  ·  ".join(str(m[k]) for k in ("source", "course", "level", "date") if m.get(k))
        if meta_line:
            self.textbox(s, MARGIN, 5.6, W - 2 * MARGIN, 0.9, [meta_line], 13, color="C9CEE6")

    def bullet_slide(self, title, items, kicker=None, notes=None, numbered=False, start=22):
        paras = []
        for it in items:
            paras.append((item_text(it), 0, {"bullet": "num" if numbered else "•"}))
            for ch in item_children(it):
                paras.append((item_text(ch), 1, {"bullet": "–"}))
        width, height = W - 2 * MARGIN, BODY_BOTTOM - BODY_TOP
        for i, (chunk, size) in enumerate(self.split_to_fit(paras, width, height, start)):
            s = self.new_slide(notes)
            self.chrome(s, title, kicker, continued=i > 0)
            self.textbox(s, MARGIN, BODY_TOP, width, height, chunk, size)

    def objectives_slide(self):
        if self.n.get("learning_objectives"):
            self.bullet_slide(self.L["objectives"], self.n["learning_objectives"], numbered=True,
                              notes="By the end of this deck you should be able to do each of these.")
        if self.n.get("prerequisites"):
            self.bullet_slide(self.L["prereq"], self.n["prerequisites"])

    def roadmap_slide(self):
        secs = self.n["sections"]
        if len(secs) < 2:
            return
        s = self.new_slide()
        self.chrome(s, self.L["agenda"])
        n = len(secs)
        cols = 2 if n > 6 else 1
        per = math.ceil(n / cols)
        colw = (W - 2 * MARGIN - 0.4 * (cols - 1)) / cols
        rowh = min(0.78, (BODY_BOTTOM - BODY_TOP) / per)
        for i, sec in enumerate(secs):
            c, r = divmod(i, per)
            x = MARGIN + c * (colw + 0.4)
            if self.rtl:
                x = W - MARGIN - colw - c * (colw + 0.4)
            y = BODY_TOP + r * rowh
            badge_x = x + colw - 0.55 if self.rtl else x
            self.rect(s, badge_x, y + 0.08, 0.55, rowh - 0.2, self.T["primary"], shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                      radius=0.2)
            self.textbox(s, badge_x, y + 0.08, 0.55, rowh - 0.2, [str(i + 1)], 16, color="FFFFFF", bold=True,
                         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
            tx = x if self.rtl else x + 0.7
            size, _ = self.fit_size([plain(sec.get("title", ""))], colw - 0.8, rowh - 0.15, 18, 12)
            self.textbox(s, tx, y + 0.04, colw - 0.75, rowh - 0.1, [sec.get("title", "")], size,
                         anchor=MSO_ANCHOR.MIDDLE)

    def divider(self, num, sec):
        s = self.new_slide(dark=True, notes=sec.get("summary"))
        self.k += 1
        self.textbox(s, MARGIN, 1.6, 3, 1.6, [num], 88, color=self.T["accent"], bold=True, font=self.hfont)
        size, _ = self.fit_size([plain(sec.get("title", ""))], W - 2 * MARGIN, 1.5, 40, 26)
        self.textbox(s, MARGIN, 3.25, W - 2 * MARGIN, 1.5, [sec.get("title", "")], size, color="FFFFFF",
                     bold=True, font=self.hfont)
        if sec.get("summary"):
            ssize, _ = self.fit_size([sec["summary"]], W - 2 * MARGIN - 1.5, 1.9, 18, 13)
            self.textbox(s, MARGIN, 4.85, W - 2 * MARGIN - 1.5, 1.9, [sec["summary"]], ssize, color="DDE1F2")
        if sec.get("source_ref"):
            self.textbox(s, MARGIN, H - 0.6, W - 2 * MARGIN, 0.35, [f"{self.L['source']}: {sec['source_ref']}"],
                         10, color="AEB4D6")

    def definition_slide(self, title, defs, kicker):
        longest = max((len(plain(d.get("definition", d.get("text", "")))) + len(plain(d.get("example", ""))))
                      for d in defs) if defs else 0
        per = 6 if longest < 170 else 4 if longest < 280 else 2
        for pg in range(0, len(defs), per):
            chunk = defs[pg:pg + per]
            s = self.new_slide("\n".join(f"{plain(d.get('term', ''))}: {plain(d.get('definition', ''))} "
                                         f"{plain(d.get('note', ''))}" for d in chunk))
            self.chrome(s, title, kicker, continued=pg > 0)
            cols = 3 if per == 6 and len(chunk) > 4 else (2 if len(chunk) > 1 else 1)
            rows = math.ceil(len(chunk) / cols)
            gw = (W - 2 * MARGIN - 0.3 * (cols - 1)) / cols
            gh = (BODY_BOTTOM - BODY_TOP - 0.25 * (rows - 1)) / rows
            need = max(self.est_height([d.get("definition", d.get("text", ""))] +
                                       ([f"e.g. {d['example']}"] if d.get("example") else []), gw - 0.5, 20) + 1.1
                       for d in chunk)
            gh = min(gh, max(need, 1.6))
            for i, d in enumerate(chunk):
                r, c = divmod(i, cols)
                if self.rtl:
                    c = cols - 1 - c
                x, y = MARGIN + c * (gw + 0.3), BODY_TOP + r * (gh + 0.25)
                self.rect(s, x, y, gw, gh, self.T["soft"], shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
                self.rect(s, x + gw - 0.09 if self.rtl else x, y + 0.15, 0.09, gh - 0.3, self.T["primary"])
                paras = [(d.get("term", ""), 0, {"bold": True, "color": self.T["primary"], "size": 0})]
                body = [(d.get("definition", d.get("text", "")), 0, {})]
                if d.get("example"):
                    body.append((f"*e.g.* {d['example']}", 0, {"color": self.T["muted"]}))
                inner_w = gw - 0.5
                size, _ = self.fit_size(body, inner_w, gh - 0.95, 20)
                paras[0][2]["size"] = size + 3
                self.textbox(s, x + 0.25, y + 0.18, inner_w, gh - 0.3, paras + body, size)

    def callout_slide(self, title, blocks, kicker):
        for pg in range(0, len(blocks), 3):
            chunk = blocks[pg:pg + 3]
            s = self.new_slide(" ".join(plain(b.get("text", "")) for b in chunk))
            self.chrome(s, title, kicker, continued=pg > 0)
            n = len(chunk)
            gap = 0.22
            avail = BODY_BOTTOM - BODY_TOP - gap * (n - 1)
            bodies = []
            for b in chunk:
                body = ([(b["text"], 0, {})] if b.get("text") else []) + \
                       [(item_text(it), 0, {"bullet": "•"}) for it in b.get("items", [])]
                bodies.append(body)
            need = [self.est_height(bd, W - 2 * MARGIN - 0.6, 20) + 0.75 for bd in bodies]
            if sum(need) <= avail:
                heights = need
            else:
                heights = [avail * x / sum(need) for x in need]
            y = BODY_TOP
            for b, h in zip(chunk, heights):
                kind = {"pitfall": "warning", "mistake": "warning", "exam_tip": "tip", "important": "key_point",
                        "callout": "key_point"}.get(b["type"], b["type"])
                bg, bar, _ = CALLOUTS.get(kind, CALLOUTS["note"])
                self.rect(s, MARGIN, y, W - 2 * MARGIN, h, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
                self.rect(s, W - MARGIN - 0.1 if self.rtl else MARGIN, y + 0.12, 0.1, h - 0.24, bar)
                label = b.get("title") or self.L.get(kind, kind)
                paras = [(label.upper() if not self.rtl else label, 0, {"bold": True, "color": bar, "size": 12})]
                body = []
                if b.get("text"):
                    body.append((b["text"], 0, {}))
                for it in b.get("items", []):
                    body.append((item_text(it), 0, {"bullet": "•", "bullet_color": bar}))
                size, _ = self.fit_size(body, W - 2 * MARGIN - 0.6, h - 0.6, 20)
                self.textbox(s, MARGIN + 0.3, y + 0.12, W - 2 * MARGIN - 0.6, h - 0.2, paras + body, size)
                y += h + gap

    def example_slide(self, b, kicker):
        title = f"{self.L['example']}: {plain(b['title'])}" if b.get("title") else self.L["example"]
        steps = [item_text(x) for x in b.get("steps", [])]
        notes = " ".join(filter(None, [plain(b.get("problem", "")), plain(b.get("text", "")),
                                      " ".join(plain(x) for x in steps), plain(b.get("solution", "")),
                                      plain(b.get("answer", ""))]))
        left = []
        if b.get("problem"):
            left.append((self.L["problem"], 0, {"bold": True, "color": self.T["primary"], "size": 13}))
            left.append((b["problem"], 0, {}))
        if b.get("text"):
            left.append((b["text"], 0, {}))
        if b.get("code"):
            left.append((b["code"], 0, {"font": self.code_font, "size": 13}))
        right = [(st, 0, {"bullet": "num"}) for st in steps]
        if b.get("solution"):
            right.append((b["solution"], 0, {}))
        if b.get("answer"):
            right.append((f"∴ **{b['answer']}**", 0, {"color": self.T["primary"]}))
        if not left:
            left, right = right, []
        two_col = bool(right) and bool(left)
        lw = (W - 2 * MARGIN) * (0.42 if two_col else 1.0)
        rw = (W - 2 * MARGIN) - lw - 0.3
        height = BODY_BOTTOM - BODY_TOP
        lsize, lfit = self.fit_size(left, lw - 0.5, height - 0.4, 20)
        rpages = self.split_to_fit(right, rw - 0.4, height - 0.4, 20) if right else [([], 20)]
        for i, (rchunk, rsize) in enumerate(rpages):
            s = self.new_slide(notes if i == 0 else None)
            self.chrome(s, title, kicker, continued=i > 0)
            lx = W - MARGIN - lw if self.rtl else MARGIN
            rx = MARGIN if self.rtl else MARGIN + lw + 0.3
            box_h = min(height, max(self.est_height(left, lw - 0.5, lsize),
                                    self.est_height(rchunk, rw - 0.4, rsize) if two_col else 0) + 0.75)
            self.rect(s, lx, BODY_TOP, lw, box_h, self.T["soft"], shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
            self.textbox(s, lx + 0.25, BODY_TOP + 0.2, lw - 0.5, height - 0.4, left, lsize)
            if two_col:
                self.textbox(s, rx + 0.1, BODY_TOP + 0.2, rw - 0.2, height - 0.4,
                             [(self.L["solution"], 0, {"bold": True, "color": self.T["primary"], "size": 13})]
                             + rchunk, rsize)

    def formula_slide(self, b, kicker):
        title = b.get("name") or self.L.get("formula", "Formula")
        s = self.new_slide(" ".join(filter(None, [plain(b.get("explanation", "")),
                                                  "; ".join(f"{v.get('symbol')}: {v.get('meaning')}"
                                                            for v in b.get("variables", []))])))
        self.chrome(s, title, kicker)
        y = BODY_TOP + 0.1
        box_h = 1.7
        self.rect(s, MARGIN, y, W - 2 * MARGIN, box_h, self.T["soft"], shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                  radius=0.08)
        png = self.tmp / f"f{self.k}_{id(b)}.png"
        if b.get("latex") and render_latex(b["latex"], png, fontsize=30, color=self.T["primary"]):
            w, h = image_size(png)
            ih = min(box_h - 0.4, h / 220)
            iw = ih * w / h
            if iw > W - 2 * MARGIN - 0.6:
                iw = W - 2 * MARGIN - 0.6
                ih = iw * h / w
            s.shapes.add_picture(str(png), Inches((W - iw) / 2), Inches(y + (box_h - ih) / 2), Inches(iw), Inches(ih))
        else:
            self.textbox(s, MARGIN, y, W - 2 * MARGIN, box_h, [b.get("text") or b.get("formula") or b.get("latex", "")],
                         30, color=self.T["primary"], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                         font="Cambria Math")
        paras = []
        if b.get("variables"):
            for v in b["variables"]:
                paras.append((f"**{v.get('symbol', '')}** — {v.get('meaning', '')}", 0, {"bullet": "•"}))
        if b.get("explanation"):
            paras.append((b["explanation"], 0, {"color": self.T["muted"]}))
        if paras:
            top = y + box_h + 0.3
            size, _ = self.fit_size(paras, W - 2 * MARGIN, BODY_BOTTOM - top, 20)
            self.textbox(s, MARGIN, top, W - 2 * MARGIN, BODY_BOTTOM - top, paras, size)

    def code_slide(self, b, kicker):
        lines = b.get("code", "").rstrip("\n").split("\n")
        height = BODY_BOTTOM - BODY_TOP
        has_expl = bool(b.get("explanation"))
        code_h = height - (1.0 if has_expl else 0)
        longest = max((len(ln) for ln in lines), default=10)
        size = 16
        while size > 10 and (len(lines) * size * 1.25 / 72 > code_h - 0.4 or
                             longest * size * 0.6 / 72 > W - 2 * MARGIN - 0.5):
            size -= 1
        per = max(4, int((code_h - 0.4) * 72 / (size * 1.25)))
        title = b.get("title") or b.get("caption") or (b.get("language") or self.L["code"]).title()
        for pg in range(0, len(lines), per):
            chunk = lines[pg:pg + per]
            s = self.new_slide(b.get("explanation"))
            self.chrome(s, title, kicker, continued=pg > 0)
            box_h = min(code_h, len(chunk) * size * 1.25 / 72 + 0.5)
            self.rect(s, MARGIN, BODY_TOP, W - 2 * MARGIN, box_h, "1E2230", shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                      radius=0.03)
            tb = self.textbox(s, MARGIN + 0.25, BODY_TOP + 0.2, W - 2 * MARGIN - 0.5, code_h - 0.4,
                              [(ln if ln else " ", 0, {"space_after": 0, "font": self.code_font,
                                                       "align": PP_ALIGN.LEFT}) for ln in chunk],
                              size, color="E6E8EF", spacing=1.0)
            for p in tb.text_frame.paragraphs:
                p._p.get_or_add_pPr().set("rtl", "0")
            if has_expl and pg == 0:
                esize, _ = self.fit_size([b["explanation"]], W - 2 * MARGIN, 0.85, 16, 11)
                self.textbox(s, MARGIN, BODY_TOP + box_h + 0.15, W - 2 * MARGIN, 0.85, [b["explanation"]], esize,
                             color=self.T["muted"])

    def table_slide(self, b, kicker, title=None):
        headers = b.get("headers") or []
        rows = b.get("rows") or []
        ncols = max([len(headers)] + [len(r) for r in rows] + [1])
        title = title or b.get("title") or b.get("caption") or "Comparison"
        width = W - 2 * MARGIN
        avail = BODY_BOTTOM - BODY_TOP - 0.1
        # choose font size & rows per slide
        col_chars = [max([len(plain(str(headers[i]))) if i < len(headers) else 0] +
                         [len(plain(str(r[i]))) if i < len(r) else 0 for r in rows] + [4]) for i in range(ncols)]
        tot = sum(min(c, 60) for c in col_chars)
        col_w = [width * min(c, 60) / tot for c in col_chars]
        col_w = [max(w, 1.1) for w in col_w]
        scale = width / sum(col_w)
        col_w = [w * scale for w in col_w]

        def row_h(r, size):
            lines = 1
            for i in range(ncols):
                txt = plain(str(r[i])) if i < len(r) else ""
                cpl = max(4, (col_w[i] - 0.2) * 72 / (size * 0.5))
                lines = max(lines, math.ceil(max(1, len(txt)) / cpl))
            return lines * size * 1.2 / 72 + 0.16

        size = 16
        while size > 11 and (row_h(headers, size) if headers else 0) + sum(row_h(r, size) for r in rows) > avail * 1.6:
            size -= 1
        pages, cur, used = [], [], row_h(headers, size) if headers else 0
        for r in rows:
            h = row_h(r, size)
            if cur and used + h > avail:
                pages.append(cur)
                cur, used = [], (row_h(headers, size) if headers else 0)
            cur.append(r)
            used += h
        pages.append(cur)
        for pi, pr in enumerate(pages):
            s = self.new_slide(b.get("note"))
            self.chrome(s, title, kicker, continued=pi > 0)
            nrows = len(pr) + (1 if headers else 0)
            total_h = (row_h(headers, size) if headers else 0) + sum(row_h(r, size) for r in pr)
            gt = s.shapes.add_table(nrows, ncols, Inches(MARGIN), Inches(BODY_TOP), Inches(width),
                                    Inches(min(total_h, avail))).table
            tblpr = gt._tbl.tblPr
            if self.rtl:
                tblpr.set("rtl", "1")
            for i, w in enumerate(col_w):
                gt.columns[i].width = Inches(w)
            all_rows = ([headers] if headers else []) + pr
            for ri, r in enumerate(all_rows):
                is_h = headers and ri == 0
                gt.rows[ri].height = Inches(row_h(r, size))
                for ci in range(ncols):
                    cell = gt.cell(ri, ci)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = rgb(self.T["primary"] if is_h else
                                                   (self.T["soft"] if ri % 2 == 0 else "FFFFFF"))
                    cell.margin_left = cell.margin_right = Inches(0.08)
                    cell.margin_top = cell.margin_bottom = Inches(0.04)
                    tf = cell.text_frame
                    tf.word_wrap = True
                    p = tf.paragraphs[0]
                    self._para_dir(p)
                    txt = str(r[ci]) if ci < len(r) else ""
                    self._runs(p, txt, size, "FFFFFF" if is_h else self.T["ink"],
                               bold=bool(is_h or (ci == 0 and ncols > 2)))

    def image_slide(self, path: Path, title, caption, kicker, notes=None):
        s = self.new_slide(notes or caption)
        self.chrome(s, title or caption or "Figure", kicker)
        cap_h = 0.45 if caption else 0
        bw, bh = W - 2 * MARGIN, BODY_BOTTOM - BODY_TOP - cap_h
        w, h = image_size(path)
        scale = min(bw / w, bh / h)
        iw, ih = w * scale, h * scale
        s.shapes.add_picture(str(path), Inches((W - iw) / 2), Inches(BODY_TOP + (bh - ih) / 2), Inches(iw), Inches(ih))
        if caption:
            self.textbox(s, MARGIN, BODY_BOTTOM - cap_h + 0.05, bw, cap_h, [caption], 12, color=self.T["muted"],
                         italic=True, align=PP_ALIGN.CENTER)

    def quiz_slides(self, qs, title, kicker=None, answer_slide=True):
        if not qs:
            return
        start = self.qn
        qparas = []
        for q in qs:
            self.qn += 1
            qparas.append((f"**Q{self.qn}.** {q.get('question', '')}", 0, {}))
            for oi, opt in enumerate(q.get("options") or []):
                qparas.append((f"{chr(65 + oi)}. {opt}", 1, {"size": 0}))
        # fix nested option sizes later in fit (size=0 sentinel -> use size-2)
        qparas = [(t, l, {k: v for k, v in o.items() if not (k == "size" and v == 0)}) for t, l, o in qparas]
        answers_txt = " ".join(f"Q{i}. {plain(q.get('answer', ''))} {plain(q.get('explanation', ''))}"
                               for i, q in enumerate(qs, start + 1))
        if not answer_slide:
            self.bullet_slide_raw(title, qparas, kicker,
                                  notes=f"{self.L['answers']}: {answers_txt}")
            return
        self.bullet_slide_raw(title, qparas, kicker, notes="Pause here and answer before moving on. "
                              f"{self.L['answers']}: {answers_txt}")
        aparas = []
        for i, q in enumerate(qs, start + 1):
            aparas.append((f"**Q{i}.** {q.get('answer', '')}", 0, {}))
            if q.get("explanation"):
                aparas.append((q["explanation"], 1, {"color": self.T["muted"]}))
        self.bullet_slide_raw(f"{title} — {self.L['answers']}", aparas, kicker)

    def bullet_slide_raw(self, title, paras, kicker=None, notes=None, start=20):
        width, height = W - 2 * MARGIN, BODY_BOTTOM - BODY_TOP
        for i, (chunk, size) in enumerate(self.split_to_fit(paras, width, height, start)):
            s = self.new_slide(notes)
            self.chrome(s, title, kicker, continued=i > 0)
            self.textbox(s, MARGIN, BODY_TOP, width, height, chunk, size)

    def cheat_slides(self, groups):
        groups = [g if isinstance(g, dict) else {"heading": "", "items": [g]} for g in groups]
        per = 6
        for pg in range(0, len(groups), per):
            chunk = groups[pg:pg + per]
            s = self.new_slide()
            self.chrome(s, self.L["cheat"], continued=pg > 0)
            cols = 3 if len(chunk) > 4 else 2 if len(chunk) > 1 else 1
            rows = math.ceil(len(chunk) / cols)
            gw = (W - 2 * MARGIN - 0.25 * (cols - 1)) / cols
            gh = (BODY_BOTTOM - BODY_TOP - 0.2 * (rows - 1)) / rows
            for i, g in enumerate(chunk):
                r, c = divmod(i, cols)
                if self.rtl:
                    c = cols - 1 - c
                x, y = MARGIN + c * (gw + 0.25), BODY_TOP + r * (gh + 0.2)
                self.rect(s, x, y, gw, gh, self.T["soft"], shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
                body = [(item_text(it), 0, {"bullet": "•"}) for it in g.get("items", [])]
                size, _ = self.fit_size(body, gw - 0.35, gh - 0.65, 15, 9)
                self.textbox(s, x + 0.18, y + 0.1, gw - 0.3, gh - 0.15,
                             [(g.get("heading", ""), 0, {"bold": True, "color": self.T["primary"], "size": size + 2})]
                             + body, size)

    # ------------------------------------------------------------------ section rendering
    CALLOUT_TYPES = ("key_point", "warning", "tip", "mnemonic", "note", "intuition", "pitfall", "mistake",
                     "exam_tip", "important", "callout")
    NOTES_ONLY = ("note", "intuition")  # elaborations: speaker notes in the deck, boxes in the doc

    @staticmethod
    def _norm(t):
        t = (t or "text").lower()
        return {"paragraph": "text", "list": "bullets", "steps": "numbered", "numbered_list": "numbered",
                "figure": "image", "math": "formula", "equation": "formula", "diagram_mermaid": "diagram",
                "pitfall": "warning", "mistake": "warning", "exam_tip": "tip", "important": "key_point",
                "callout": "key_point"}.get(t, t)

    def top_section(self, num, sec, quiz_pool):
        """Render one top-level section and all its subsections as a compact review sequence:
        divider → key-term cards → content (per subsection) → 'remember' slide → check-yourself quiz."""
        self.root_title = sec.get("title", "")
        self.divider(num, sec)
        subtree = [(num, 1, sec)] + [(f"{num}.{n}", d + 1, s)
                                     for n, d, s in walk_sections(sec.get("subsections", []))]
        defs = [b for _, _, s in subtree for b in s.get("blocks", [])
                if self._norm(b.get("type")) == "definition" and b.get("slide", True) is not False]
        if defs:
            self.definition_slide(self.L.get("key_terms", "Key terms"),
                                  defs, f"§{num}")
        remember = []
        self._flow = {"text": [], "notes": [], "code": [], "title": None, "kicker": None}
        for snum, depth, s in subtree:
            self.content_flow(snum, depth, s, remember)
            for t in s.get("takeaways", []) or []:
                remember.append(("key_point", item_text(t)))
        self._flush_flow()
        if remember:
            paras = []
            for kind, text in remember:
                bg, bar, _ = CALLOUTS.get(kind, CALLOUTS["note"])
                paras.append((f"**{self.L.get(kind, kind)}:** {text}", 0, {"bullet": "■", "bullet_color": bar}))
            self.bullet_slide_raw(f"{self.L.get('remember', 'Remember')}: {plain(sec.get('title', ''))}", paras,
                                  f"§{num}", notes=" ".join(plain(t) for _, t in remember))
        if quiz_pool:
            self.quiz_slides(quiz_pool, f"{self.L['check']}: {plain(sec.get('title', ''))}", f"§{num}",
                             answer_slide=bool(self.meta.get("deck_section_answer_slides", False)))

    def content_flow(self, num, depth, sec, remember):
        """Text from consecutive (sub)sections flows onto shared slides (with sub-headings) until a
        visual block needs its own slide — this keeps review decks compact."""
        title = sec.get("title", "")
        kicker = f"§{num} · {plain(self.root_title)}" if depth > 1 else f"§{num}"
        blocks = [b for b in sec.get("blocks", []) if b.get("slide", True) is not False]
        F = self._flow
        pending_text, pending_notes, pending_code = F["text"], F["notes"], F["code"]
        if F["title"] is None or not pending_text:
            F["title"], F["kicker"] = title, kicker
        if depth > 1:
            pending_text.append((f"{num}  {title}", 0, {"bold": True, "color": self.T["primary"],
                                                         "space_after": 4}))
            if sec.get("summary"):
                pending_text.append((sec["summary"], 0, {"italic": True, "color": self.T["muted"]}))

        def flush_text():
            if pending_text:
                ft = F["title"] or title
                paras = list(pending_text)
                if paras and paras[0][2].get("bold") and paras[0][0].endswith(plain(ft)):
                    paras = paras[1:]  # sub-heading duplicates the slide title
                if paras:
                    self.bullet_slide_raw(ft, paras, F["kicker"] or kicker, notes=" ".join(pending_notes))
                pending_text.clear()
                pending_notes.clear()
            F["title"], F["kicker"] = title, kicker

        def flush_code():
            while pending_code:
                a = pending_code.pop(0)
                if pending_code and self._small_code(a) and self._small_code(pending_code[0]):
                    self.code_pair_slide(a, pending_code.pop(0), title, kicker)
                else:
                    self.code_slide(a, kicker)

        def flush():
            flush_text()
            flush_code()

        self._flush_flow = flush
        for i, b in enumerate(blocks):
            t = self._norm(b.get("type"))
            if t == "text":
                short = b.get("short") or b.get("slide_text")
                txt = b.get("text", "")
                if short:
                    for x in (short if isinstance(short, list) else [short]):
                        pending_text.append((x, 0, {"bullet": "•"}))
                else:
                    for sent in sentences(txt)[:4]:
                        pending_text.append((sent, 0, {"bullet": "•"}))
                pending_notes.append(plain(txt))
            elif t in ("bullets", "numbered"):
                if b.get("title"):
                    pending_text.append((b["title"], 0, {"bold": True, "color": self.T["primary"]}))
                for it in b.get("items", []):
                    pending_text.append((item_text(it), 0, {"bullet": "num" if t == "numbered" else "•"}))
                    for ch in item_children(it):
                        pending_text.append((item_text(ch), 1, {"bullet": "–"}))
            elif t == "definition":
                continue  # shown on the section's key-term cards
            elif t in self.NOTES_ONLY:
                pending_notes.append(plain(" ".join(filter(None, [b.get("title"), b.get("text")] +
                                                          [item_text(x) for x in b.get("items", [])]))))
            elif t in ("key_point", "warning", "tip", "mnemonic"):
                if b.get("text"):
                    remember.append((t, b["text"]))
                for it in b.get("items", []):
                    remember.append((t, item_text(it)))
            elif t == "code":
                pending_code.append(b)
            elif t == "example":
                flush()
                self.example_slide(b, kicker)
            elif t == "formula":
                flush()
                self.formula_slide(b, kicker)
            elif t == "table":
                flush()
                self.table_slide(b, kicker, b.get("title") or b.get("caption") or title)
            elif t == "image":
                flush()
                p = Path(b.get("path", "")).expanduser()
                p = p if p.is_absolute() else (self.base_dir / p)
                if p.exists():
                    self.image_slide(p, b.get("title") or title, b.get("caption"), kicker)
            elif t == "diagram":
                flush()
                png = self.tmp / f"d{self.k}_{i}.png"
                code = b.get("mermaid") or b.get("code") or ""
                if code and render_mermaid(code, png, width=2000):
                    self.image_slide(png, b.get("title") or title, b.get("caption"), kicker)
                elif b.get("ascii"):
                    self.code_slide({"code": b["ascii"], "title": b.get("caption") or title}, kicker)
            elif t == "quote":
                pending_text.append((f"“{b.get('text', '')}”" + (f" — {b['by']}" if b.get("by") else ""), 0,
                                     {"italic": True, "color": self.T["muted"]}))
            elif b.get("text"):
                pending_text.append((b["text"], 0, {"bullet": "•"}))
        if pending_code:
            flush()

    @staticmethod
    def _small_code(b):
        lines = b.get("code", "").rstrip("\n").split("\n")
        return len(lines) <= 14 and max((len(x) for x in lines), default=0) <= 58

    def code_pair_slide(self, a, b, title, kicker):
        s = self.new_slide(" ".join(filter(None, [plain(a.get("explanation", "")), plain(b.get("explanation", ""))])))
        self.chrome(s, title, kicker)
        colw = (W - 2 * MARGIN - 0.3) / 2
        for ci, blk in enumerate((a, b)):
            x = MARGIN + ci * (colw + 0.3)
            if self.rtl:
                x = W - MARGIN - colw - ci * (colw + 0.3)
            y = BODY_TOP
            cap = blk.get("title") or blk.get("caption")
            if cap:
                self.textbox(s, x, y, colw, 0.4, [cap], 13, color=self.T["primary"], bold=True)
                y += 0.45
            lines = blk.get("code", "").rstrip("\n").split("\n")
            size = 14
            while size > 10 and (len(lines) * size * 1.25 / 72 > 3.9 or
                                 max(len(x) for x in lines) * size * 0.6 / 72 > colw - 0.4):
                size -= 1
            box_h = len(lines) * size * 1.25 / 72 + 0.4
            self.rect(s, x, y, colw, box_h, "1E2230", shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
            tb = self.textbox(s, x + 0.2, y + 0.2, colw - 0.4, box_h - 0.3,
                              [(ln if ln else " ", 0, {"space_after": 0, "font": self.code_font,
                                                       "align": PP_ALIGN.LEFT}) for ln in lines],
                              size, color="E6E8EF", spacing=1.0)
            for p in tb.text_frame.paragraphs:
                p._p.get_or_add_pPr().set("rtl", "0")
            if blk.get("explanation"):
                top = y + box_h + 0.12
                esize, _ = self.fit_size([blk["explanation"]], colw, BODY_BOTTOM - top, 15, 10)
                self.textbox(s, x, top, colw, BODY_BOTTOM - top, [blk["explanation"]], esize, color=self.T["muted"])

    def build(self, out: Path):
        n, L = self.n, self.L
        self.title_slide()
        if n.get("overview"):
            ov = n["overview"] if isinstance(n["overview"], list) else [n["overview"]]
            paras = []
            for o in ov:
                paras += [(x, 0, {"bullet": "•"}) for x in sentences(o)]
            self.bullet_slide_raw(L["overview"], paras, notes=" ".join(ov))
        self.objectives_slide()
        self.roadmap_slide()
        qs_by_sec: dict[str, list] = {}
        for q in n.get("review_questions", []):
            qs_by_sec.setdefault(str(q.get("section", "")).split(".")[0], []).append(q)
        used_q = set()
        for i, sec in enumerate(n["sections"], 1):
            pool = []
            if self.meta.get("section_quizzes", True):
                pool = [q for q in qs_by_sec.get(str(i), []) if id(q) not in used_q][:3]
                used_q.update(id(q) for q in pool)
            self.top_section(str(i), sec, pool)
        if n.get("key_takeaways"):
            self.bullet_slide(L["takeaways"], n["key_takeaways"], numbered=True)
        if n.get("common_mistakes"):
            items = []
            for m in n["common_mistakes"]:
                if isinstance(m, dict):
                    items.append({"text": f"✗ {m.get('mistake', m.get('text', ''))}",
                                  "children": [f"✓ {m.get('correction') or m.get('fix') or ''}"]
                                  if (m.get('correction') or m.get('fix')) else []})
                else:
                    items.append(m)
            self.bullet_slide(L["mistakes"], items)
        if n.get("exam_tips"):
            self.callout_slide(L["exam_tips"], [{"type": "tip", "items": n["exam_tips"], "title": L["exam_tips"]}],
                               None)
        if n.get("cheat_sheet"):
            self.cheat_slides(n["cheat_sheet"])
        if n.get("glossary"):
            items = sorted(n["glossary"], key=lambda g: plain(g.get("term", "")).lower())
            self.table_slide({"headers": [L["term"], L["definition"]],
                              "rows": [[g.get("term", ""), g.get("definition", "")] for g in items]},
                             None, L["glossary"])
        remaining = [q for q in n.get("review_questions", []) if id(q) not in used_q]
        cap = int(self.meta.get("deck_final_quiz", 12))
        if remaining and cap > 0:
            # a varied final quiz: spread picks across sections and difficulties; the doc keeps all questions
            remaining.sort(key=lambda q: ({"easy": 0, "medium": 1, "hard": 2}.get(q.get("difficulty"), 1)))
            pick = remaining[::max(1, len(remaining) // cap)][:cap]
            for i in range(0, len(pick), 6):
                self.quiz_slides(pick[i:i + 6], L["questions"])
        # every slide gets speaker notes: fall back to the slide's own text
        for sl in self.prs.slides:
            tf = sl.notes_slide.notes_text_frame
            if len(tf.text.strip()) < 20:
                body = [sh.text_frame.text for sh in sl.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
                tf.text = (tf.text + "\n" if tf.text.strip() else "") + "\n".join(body[:-2] if len(body) > 2 else body)
        out.parent.mkdir(parents=True, exist_ok=True)
        self.prs.core_properties.title = plain(self.meta.get("title", ""))
        self.prs.save(str(out))
        return out, len(self.prs.slides)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("notes")
    ap.add_argument("-o", "--output", required=True)
    a = ap.parse_args()
    notes = load_notes(a.notes)
    out, n = Deck(notes, Path(a.notes).resolve().parent).build(Path(a.output))
    print(f"[build_pptx] wrote {out} ({n} slides)")
    if n > 90:
        print("[build_pptx] NOTE: the deck is long for review. Consider marking secondary blocks with "
              '"slide": false, giving long paragraphs a short "short" bullet list, or merging small tables.')


if __name__ == "__main__":
    main()
