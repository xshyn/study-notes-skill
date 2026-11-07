#!/usr/bin/env python3
"""
validate_notes.py — check a study-notes JSON before rendering, and measure how well it
covers the source.

  python validate_notes.py notes.json [--source source.txt] [--depth standard|exam|deep]

Reports
  * ERRORS   – structural problems that would break or weaken rendering (exit code 1)
  * WARNINGS – quality gaps for learning material (thin sections, unanswered questions …)
  * COVERAGE – (with --source) source headings / key terms / pages the notes never mention,
               so you can go back and fill gaps instead of silently dropping material.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from notes_common import item_text, load_notes, plain, walk_sections  # noqa: E402

KNOWN = {"text", "paragraph", "bullets", "list", "numbered", "steps", "numbered_list", "definition", "key_point",
         "warning", "tip", "mnemonic", "note", "intuition", "pitfall", "mistake", "exam_tip", "important", "callout",
         "example", "formula", "math", "equation", "code", "table", "quote", "image", "figure", "diagram",
         "diagram_mermaid", "chart", "graph", "plot"}

TARGETS = {  # minimums per depth profile
    "summary":  {"blocks_per_section": 3, "questions_min": 6,  "q_per_section": 1, "glossary_min": 5,  "flash_min": 0},
    "standard": {"blocks_per_section": 5, "questions_min": 12, "q_per_section": 2, "glossary_min": 10, "flash_min": 10},
    "exam":     {"blocks_per_section": 6, "questions_min": 20, "q_per_section": 3, "glossary_min": 12, "flash_min": 15},
    "deep":     {"blocks_per_section": 8, "questions_min": 25, "q_per_section": 3, "glossary_min": 15, "flash_min": 20},
}

STOP = set("""a an the and or of to in on for with by is are was were be been being this that these those it its as at
from into than then there their they them which who whom whose what when where why how can could should would will
may might must not no nor so such also only other some any each every all both more most many much very one two three
we you he she his her our your i if else about over under between within without per via e g eg ie use used using
example examples figure fig table chapter section page pages see note however therefore thus""".split())


def all_text(notes: dict) -> str:
    chunks = [json.dumps(notes, ensure_ascii=False)]
    return plain(" ".join(chunks)).lower()


def terms_from_source(src: str, top: int) -> list[str]:
    heads = [re.sub(r"^#+\s*", "", ln).strip() for ln in src.splitlines() if re.match(r"^#{1,4}\s+\S", ln)]
    # key terms come from prose only: skip code-looking lines (prompts, calls, assignments, symbols)
    prose = "\n".join(ln for ln in src.splitlines()
                      if not re.match(r"^\s*(>>>|\.\.\.|\$|#!|def |class |import |from )", ln)
                      and len(re.findall(r"[(){}\[\]=<>;\\|^$*+?]", ln)) <= max(2, len(ln) // 25))
    words = re.findall(r"[A-Za-z][A-Za-z\-]{3,}", prose)
    low = [w.lower() for w in words if w.lower() not in STOP]
    uni = Counter(low)
    bi = Counter(f"{a} {b}" for a, b in zip(low, low[1:]) if a != b)
    cands = [w for w, c in bi.most_common(top * 3) if c >= 5][:top // 2] + \
            [w for w, c in uni.most_common(top * 2) if c >= 4][:top]
    return heads, list(dict.fromkeys(cands))


def page_ranges(ref: str) -> set[int]:
    pages = set()
    for a, b in re.findall(r"(\d+)\s*[–\-—]\s*(\d+)", ref or ""):
        a, b = int(a), int(b)
        if 0 < b - a < 400:
            pages.update(range(a, b + 1))
    stripped = re.sub(r"\d+\s*[–\-—]\s*\d+", " ", ref or "")
    pages.update(int(x) for x in re.findall(r"\b(\d{1,4})\b", stripped))
    return pages


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("notes")
    ap.add_argument("--source", help="extracted source text, for coverage checks")
    ap.add_argument("--depth", default=None, choices=list(TARGETS))
    ap.add_argument("--json", action="store_true", help="print machine-readable report")
    a = ap.parse_args()

    errors, warns, info = [], [], []
    try:
        n = load_notes(a.notes)
    except json.JSONDecodeError as e:
        print(f"ERROR: invalid JSON — {e}")
        sys.exit(1)
    depth = a.depth or n["meta"].get("depth", "standard")
    tgt = TARGETS.get(depth, TARGETS["standard"])

    if not n["meta"].get("title"):
        errors.append("meta.title is missing")
    if not n["sections"]:
        errors.append("no sections")

    secs = list(walk_sections(n["sections"]))
    counts = Counter()
    for num, depth_, s in secs:
        if not s.get("title"):
            errors.append(f"section {num} has no title")
        blocks = s.get("blocks", [])
        if depth_ == 1 and len(blocks) + sum(len(x.get("blocks", [])) for x in s.get("subsections", [])) \
                < tgt["blocks_per_section"]:
            warns.append(f"section {num} '{s.get('title', '')[:40]}' is thin "
                         f"({len(blocks)} blocks; target ≥ {tgt['blocks_per_section']} for depth={depth})")
        if depth_ == 1 and not s.get("summary"):
            warns.append(f"section {num} has no 'summary' (used for the In-brief box and slide divider)")
        for bi, b in enumerate(blocks):
            t = (b.get("type") or "text").lower()
            counts[t] += 1
            where = f"section {num} block {bi + 1} ({t})"
            if t not in KNOWN:
                warns.append(f"{where}: unknown type, will render as text")
            if t == "definition" and not (b.get("term") and (b.get("definition") or b.get("text"))):
                errors.append(f"{where}: definition needs 'term' and 'definition'")
            if t == "table":
                rows = b.get("rows") or []
                if not rows:
                    errors.append(f"{where}: table has no rows")
                hdr = len(b.get("headers") or [])
                if hdr and any(len(r) != hdr for r in rows):
                    warns.append(f"{where}: some rows don't match header width ({hdr})")
            if t in ("chart", "graph", "plot") and not any(x.get("values") for x in b.get("series", [])):
                errors.append(f"{where}: chart needs series with values")
            if t == "code" and not b.get("code"):
                errors.append(f"{where}: code block is empty")
            if t in ("formula", "math", "equation") and not (b.get("latex") or b.get("text") or b.get("formula")):
                errors.append(f"{where}: formula needs 'latex' or 'text'")
            if t in ("formula", "math", "equation") and b.get("latex") and not b.get("text"):
                info.append(f"{where}: tip — add a plain 'text' fallback for the formula")
            if t == "example" and not (b.get("steps") or b.get("solution") or b.get("answer") or b.get("text")):
                warns.append(f"{where}: worked example has no steps/solution")
            if t in ("image", "figure"):
                p = Path(b.get("path", ""))
                p = p if p.is_absolute() else Path(a.notes).resolve().parent / p
                if not p.exists():
                    errors.append(f"{where}: image not found: {b.get('path')}")
            if t == "text" and len(b.get("text", "")) > 900:
                warns.append(f"{where}: very long paragraph ({len(b['text'])} chars) — split it or add 'short' "
                             f"bullets for the slide version")
            if t in ("bullets", "numbered"):
                for it in b.get("items", []):
                    if len(item_text(it)) > 320:
                        warns.append(f"{where}: bullet over 320 chars — split into sub-bullets")
                        break

    qs = n.get("review_questions", [])
    if len(qs) < tgt["questions_min"]:
        warns.append(f"only {len(qs)} review questions (target ≥ {tgt['questions_min']} for depth={depth})")
    top_secs = [num for num, d, _ in secs if d == 1]
    want = tgt["q_per_section"]
    by_sec = Counter(str(q.get("section", "")).split(".")[0] for q in qs)
    thin_q = [s for s in top_secs if by_sec.get(s, 0) < want]
    if thin_q:
        warns.append(f"sections with < {want} tagged questions: {', '.join(thin_q)} "
                     f"(set 'section' on each question)")
    types = Counter(q.get("type", "?") for q in qs)
    if qs and len(types) < 3:
        warns.append(f"question types are not varied ({dict(types)}); mix recall / conceptual / application")
    for i, q in enumerate(qs, 1):
        if not q.get("question") or not q.get("answer"):
            errors.append(f"review question {i} missing question or answer")
        if q.get("type") == "mcq" and not q.get("options"):
            errors.append(f"review question {i} is mcq but has no options")
    if len(n.get("glossary", [])) < tgt["glossary_min"]:
        warns.append(f"glossary has {len(n.get('glossary', []))} terms (target ≥ {tgt['glossary_min']})")
    if len(n.get("flashcards", [])) < tgt["flash_min"]:
        warns.append(f"{len(n.get('flashcards', []))} flashcards (target ≥ {tgt['flash_min']})")
    for key in ("learning_objectives", "key_takeaways", "common_mistakes", "exam_tips", "cheat_sheet", "overview"):
        if not n.get(key):
            warns.append(f"'{key}' is empty")
    allsecs = [sec for _, _, sec in secs]
    marked = sum(1 for sec in allsecs for b in sec.get("blocks", []) if b.get("slide") is True)
    if not marked and not any(sec.get("slide_bullets") for sec in allsecs):
        warns.append("deck is not curated: no section has 'slide_bullets' and no block has \"slide\": true "
                     "— the deck will show every block and be long")
    else:
        nob = [num for num, d, sec in secs if d == 1 and not sec.get("slide_bullets")]
        if nob:
            warns.append(f"top-level sections without slide_bullets: {', '.join(nob)}")
        if marked > 12 * max(1, len(top_secs)):
            warns.append(f"{marked} blocks marked for slides — likely a long deck; keep only what earns a slide")
    if counts["example"] == 0:
        warns.append("no worked examples — learners need at least one per major concept")
    if counts["definition"] == 0:
        warns.append("no definition blocks")

    coverage = {}
    if a.source:
        src = Path(a.source).read_text(encoding="utf8", errors="ignore")
        txt = all_text(n)
        heads, keyterms = terms_from_source(src, 25)
        def _hwords(h):
            ws = [w for w in re.findall(r"[\w\-]+", h) if len(w) >= 4 or (len(w) >= 2 and w.isupper())]
            return [w.lower() for w in ws if w.lower() not in STOP][:3]
        miss_h = [h for h in heads if len(h) > 2 and _hwords(h) and not any(w in txt for w in _hwords(h))]
        miss_t = [t for t in keyterms if t not in txt]
        src_pages = sorted({int(x) for x in re.findall(r"^=== (?:Page|Slide) (\d+) ===", src, flags=re.M)})
        covered = set()
        for _, _, s in secs:
            covered |= page_ranges(s.get("source_ref", ""))
        uncovered = [p for p in src_pages if p not in covered] if covered else []
        src_words = len(src.split())
        note_words = len(txt.split())
        coverage = {"source_words": src_words, "notes_words": note_words,
                    "ratio": round(note_words / max(src_words, 1), 3),
                    "missing_headings": miss_h[:30], "missing_key_terms": miss_t[:30],
                    "key_term_coverage": round(1 - len(miss_t) / max(len(keyterms), 1), 3),
                    "uncovered_pages": uncovered}
        if miss_h:
            warns.append(f"{len(miss_h)} source headings not reflected in notes: {miss_h[:8]}")
        if keyterms and coverage["key_term_coverage"] < 0.8:
            warns.append(f"key-term coverage {coverage['key_term_coverage']:.0%}; missing: {miss_t[:12]}")
        if uncovered and len(uncovered) > max(2, 0.1 * len(src_pages)):
            warns.append(f"{len(uncovered)} source pages not cited by any section.source_ref: "
                         f"{uncovered[:20]}{'…' if len(uncovered) > 20 else ''}")

    stats = {"sections": len(top_secs), "subsections": len(secs) - len(top_secs),
             "blocks": sum(counts.values()), "block_types": dict(counts), "questions": len(qs),
             "question_types": dict(types), "glossary": len(n.get("glossary", [])),
             "flashcards": len(n.get("flashcards", [])), "depth": depth}
    if a.json:
        print(json.dumps({"errors": errors, "warnings": warns, "info": info, "stats": stats,
                          "coverage": coverage}, ensure_ascii=False, indent=2))
    else:
        print(f"STATS  {json.dumps(stats, ensure_ascii=False)}")
        if coverage:
            print(f"COVERAGE  {json.dumps(coverage, ensure_ascii=False)}")
        for e in errors:
            print("ERROR  ", e)
        for w in warns:
            print("WARN   ", w)
        for i in info[:10]:
            print("INFO   ", i)
        print("RESULT ", "FAIL" if errors else ("PASS with warnings" if warns else "PASS"))
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
