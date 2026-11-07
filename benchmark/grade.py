#!/usr/bin/env python3
"""Programmatic grader for study-notes evals. Same checks for with_skill and without_skill runs.

usage: grade.py <iteration-dir>
Writes grading.json into every run directory (<case>/<config>/).
"""
import json
import math
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
from pptx import Presentation
from pptx.util import Emu

CASE_TERMS = {
    "pdf-regex-howto": ["metacharacter", "character class", "backslash", "raw string", "compile", "match(",
                        "search(", "findall", "finditer", "greedy", "non-greedy", "group(", "named group",
                        "non-capturing", "lookahead", "negative lookahead", "sub(", "split(", "IGNORECASE",
                        "VERBOSE", "MULTILINE", "DOTALL", "\\b", "anchor", "backtrack", "{m,n}", "alternation"],
    "url-tcp-congestion": ["congestion window", "cwnd", "slow start", "ssthresh", "congestion avoidance",
                           "AIMD", "additive increase", "multiplicative decrease", "fast retransmit",
                           "fast recovery", "duplicate ACK", "timeout", "Tahoe", "Reno", "NewReno", "CUBIC",
                           "BBR", "Vegas", "ECN", "RTT", "receive window", "MSS", "congestion collapse",
                           "fairness", "packet loss"],
    "fa-exam-python-classes": ["namespace", "scope", "global", "nonlocal", "class object", "instance",
                               "method", "self", "__init__", "class variable", "instance variable",
                               "inheritance", "super", "multiple inheritance", "MRO", "private", "name mangling",
                               "iterator", "__iter__", "__next__", "StopIteration", "generator", "yield",
                               "generator expression", "isinstance", "issubclass", "dataclass"],
}
MIN_WORDS = {"pdf-regex-howto": 6000, "url-tcp-congestion": 6000, "fa-exam-python-classes": 6000}

QWORDS = re.compile(r"(question|quiz|self-test|practice|سؤال|سوال|پرسش|خودآزمایی|تمرین)", re.I)
AWORDS = re.compile(r"(answer|solution|پاسخ|جواب)", re.I)
GLOSS = re.compile(r"(glossary|key terms|vocabulary|واژه‌نامه|واژه نامه|اصطلاحات|واژگان)", re.I)
MIST = re.compile(r"(mistake|pitfall|gotcha|misconception|اشتباه|خطاهای رایج|دام)", re.I)
TIPS = re.compile(r"(exam tip|exam strateg|test tip|likely to be asked|نکات امتحان|نکته امتحانی|نکات کنکوری|توصیه.*امتحان)", re.I)
FLASH = re.compile(r"(flash ?card|فلش|کارت‌های مرور|کارت مرور)", re.I)
CHEAT = re.compile(r"(cheat ?sheet|one-page summary|quick reference|برگه خلاصه|برگهٔ خلاصه|خلاصه یک|تقلب)", re.I)
EXAMPLE = re.compile(r"(example|مثال)", re.I)


def docx_info(path):
    d = Document(str(path))
    paras = [p for p in d.paragraphs]
    texts = [p.text for p in paras]
    for t in d.tables:
        for r in t.rows:
            for c in r.cells:
                texts.append(c.text)
    full = "\n".join(texts)
    heads = [p.text for p in paras if p.style is not None and p.style.name.lower().startswith(("heading", "title"))]
    bidi = sum(1 for p in d.element.body.iter() if p.tag.endswith("}bidi"))
    table_rows = sum(len(t.rows) for t in d.tables)
    imgs = len(d.inline_shapes)
    return {"text": full, "headings": heads, "words": len(full.split()), "bidi": bidi, "paras": len(paras),
            "tables": len(d.tables), "table_rows": table_rows, "images": imgs}


def pptx_info(path):
    prs = Presentation(str(path))
    W, H = prs.slide_width, prs.slide_height
    slides, notes, overflow, out_of_bounds, sparse = [], 0, [], 0, []
    for si, s in enumerate(prs.slides, 1):
        t = []
        visual = any((getattr(sh, "has_table", False) and sh.has_table) or sh.shape_type in (13, 6) for sh in s.shapes)
        boxes = sorted([sh.text_frame.text.strip() for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()],
                       key=len, reverse=True)
        body_chars = sum(len(x) for x in boxes[1:]) if len(boxes) > 1 else 0  # minus the longest? use total minus title
        total = sum(len(x) for x in boxes)
        if not visual and total < 140 and si > 1:
            sparse.append(si)
        for shp in s.shapes:
            if shp.left is not None and shp.width is not None:
                if shp.left < -Emu(10000) or shp.top < -Emu(10000) or shp.left + shp.width > W + Emu(20000) \
                        or shp.top + shp.height > H + Emu(20000):
                    out_of_bounds += 1
            if shp.has_text_frame:
                t.append(shp.text_frame.text)
                # overflow heuristic, identical for all runs
                w = Emu(shp.width).inches - 0.15
                h = Emu(shp.height).inches
                need = 0.0
                for p in shp.text_frame.paragraphs:
                    size = None
                    for r in p.runs:
                        if r.font.size:
                            size = r.font.size.pt
                            break
                    size = size or 18
                    txt = "".join(r.text for r in p.runs)
                    if not txt.strip():
                        continue
                    cpl = max(6, w * 72 / (size * 0.5))
                    need += max(1, math.ceil(len(txt) / cpl)) * size * 1.25 / 72
                if need > h * 1.3 and need - h > 0.4 and w > 1:
                    overflow.append(si)
            if getattr(shp, "has_table", False) and shp.has_table:
                for r in shp.table.rows:
                    for c in r.cells:
                        t.append(c.text)
        if s.has_notes_slide and len(s.notes_slide.notes_text_frame.text.strip()) > 20:
            notes += 1
        slides.append("\n".join(t))
    return {"n": len(slides), "text": "\n".join(slides), "notes": notes, "overflow": sorted(set(overflow)),
            "oob": out_of_bounds, "sparse": sparse}


def opens(path):
    soffice = shutil.which("soffice")
    with tempfile.TemporaryDirectory() as td:
        r = subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", td, str(path)],
                           capture_output=True, timeout=300)
        return (Path(td) / (path.stem + ".pdf")).exists()


def distinct_q(text):
    nums = set(re.findall(r"(?m)(?:^|\s)(?:Q|س|سؤال\s*|سوال\s*)(\d{1,3})[\.\):]", text))
    qm = len([ln for ln in text.splitlines() if ln.strip().endswith(("?", "؟")) and len(ln.strip()) > 15])
    return max(len(nums), qm)


def count_questions(text):
    # numbered questions like "Q12." or lines ending with "?" / "؟"
    qn = len(re.findall(r"(?m)(^|\s)(Q|س|سؤال\s*|سوال\s*)\d{1,3}[\.\):]", text))
    qm = len([ln for ln in text.splitlines() if ln.strip().endswith(("?", "؟")) and len(ln.strip()) > 15])
    return max(qn, qm)


def grade_run(case, run_dir):
    out = run_dir / "outputs"
    docs = sorted(out.glob("*.docx"))
    decks = sorted(out.glob("*.pptx"))
    ex = []

    def add(text, ok, ev):
        ex.append({"text": text, "passed": bool(ok), "evidence": ev})

    add("A .docx study document was produced and opens cleanly (LibreOffice converts it)",
        docs and opens(docs[0]), f"docx files: {[d.name for d in docs]}")
    add("A .pptx deck was produced and opens cleanly (LibreOffice converts it)",
        decks and opens(decks[0]), f"pptx files: {[d.name for d in decks]}")
    di = docx_info(docs[0]) if docs else None
    pi = pptx_info(decks[0]) if decks else None
    if di:
        add(f"The document is detailed: at least {MIN_WORDS[case]} words", di["words"] >= MIN_WORDS[case],
            f"{di['words']} words")
        add("The document is well structured: at least 20 headings", len(di["headings"]) >= 20,
            f"{len(di['headings'])} headings")
        terms = CASE_TERMS[case]
        low = di["text"].lower()
        hit = [t for t in terms if t.lower() in low]
        add(f"The document covers ≥ 95% of the source's key concepts ({len(terms)} checked)",
            len(hit) / len(terms) >= 0.95,
            f"{len(hit)}/{len(terms)} present; missing: {[t for t in terms if t not in hit]}")
        nq = distinct_q(di["text"])
        add("The document has at least 20 practice questions (enough to self-test for an exam)", nq >= 20,
            f"~{nq} distinct questions detected")
        qh = [i for i, h in enumerate(di["headings"]) if QWORDS.search(h)]
        ah = [i for i, h in enumerate(di["headings"]) if AWORDS.search(h) and not QWORDS.search(h)]
        add("Answers are in a separate answer-key section after the questions (so the learner can self-test)",
            bool(qh and ah and max(ah) > min(qh)),
            f"question headings at {qh[:3]}, answer-only headings at {ah[:3]}")
        fl = [h for h in di["headings"] if FLASH.search(h)]
        add("The document includes flashcards for spaced review", bool(fl), f"{fl[:2]}")
        ch = [h for h in di["headings"] if CHEAT.search(h)]
        add("The document includes a cheat sheet / one-page summary", bool(ch), f"{ch[:2]}")
        g = [h for h in di["headings"] if GLOSS.search(h)]
        add("The document has a glossary / key-terms section", bool(g), f"{g[:2]}")
        m = [h for h in di["headings"] if MIST.search(h)] or re.findall(MIST, di["text"])[:3]
        add("The document flags common mistakes / pitfalls", bool(m), f"{m[:3]}")
        t_ = [h for h in di["headings"] if TIPS.search(h)] or re.findall(TIPS, di["text"])[:3]
        add("The document gives exam tips", bool(t_), f"{t_[:3]}")
        nex = len(re.findall(EXAMPLE, di["text"]))
        add("The document contains worked examples (≥ 4 mentions of example)", nex >= 4, f"{nex} mentions")
        add("The document uses tables (comparison / summary) — at least 2", di["tables"] >= 2,
            f"{di['tables']} tables")
    if pi:
        add("The deck is a reviewable size for flipping through: 15–45 slides", 15 <= pi["n"] <= 45,
            f"{pi['n']} slides")
        add("At least 90% of slides have speaker notes", pi["n"] and pi["notes"] / pi["n"] >= 0.9,
            f"{pi['notes']}/{pi['n']} slides with notes")
        add("No slide has text that overflows its box or shapes outside the slide",
            not pi["overflow"] and pi["oob"] == 0,
            f"overflow on slides {pi['overflow'][:15]}; out-of-bounds shapes: {pi['oob']}")
        add("The deck includes practice/quiz questions", count_questions(pi["text"]) >= 5 or bool(QWORDS.search(pi["text"])),
            f"~{count_questions(pi['text'])} questions on slides")
    if case == "fa-exam-python-classes" and di:
        fa_chars = len(re.findall(r"[؀-ۿ]", di["text"]))
        lat = len(re.findall(r"[A-Za-z]", di["text"]))
        add("The document is written in Persian (Persian script dominates the text)", fa_chars > lat,
            f"{fa_chars} Persian vs {lat} Latin letters")
        add("The document is laid out right-to-left (bidi paragraphs)", di["bidi"] >= di["paras"] * 0.5,
            f"{di['bidi']} bidi markers / {di['paras']} paragraphs")
        if pi:
            rtl = pi_rtl(decks[0])
            add("The deck text is laid out right-to-left", rtl >= 0.5, f"{rtl:.0%} of Persian paragraphs marked rtl")
    passed = sum(e["passed"] for e in ex)
    res = {"expectations": ex, "summary": {"passed": passed, "failed": len(ex) - passed, "total": len(ex),
                                           "pass_rate": round(passed / len(ex), 3) if ex else 0}}
    (run_dir / "grading.json").write_text(json.dumps(res, ensure_ascii=False, indent=2))
    return res


def pi_rtl(path):
    prs = Presentation(str(path))
    tot = rtl = 0
    for s in prs.slides:
        for shp in s.shapes:
            if shp.has_text_frame:
                for p in shp.text_frame.paragraphs:
                    txt = "".join(r.text for r in p.runs)
                    if re.search(r"[؀-ۿ]", txt):
                        tot += 1
                        ppr = p._p.pPr
                        if ppr is not None and ppr.get("rtl") == "1":
                            rtl += 1
    return rtl / tot if tot else 0


if __name__ == "__main__":
    it = Path(sys.argv[1])
    for case_dir in sorted(p for p in it.iterdir() if p.is_dir()):
        for cfg in ("with_skill", "without_skill"):
            rd = case_dir / cfg
            if (rd / "outputs").exists():
                r = grade_run(case_dir.name, rd)
                print(f"{case_dir.name:28s} {cfg:14s} {r['summary']['passed']}/{r['summary']['total']}")
                for e in r["expectations"]:
                    if not e["passed"]:
                        print("     FAIL", e["text"], "|", e["evidence"][:150])
