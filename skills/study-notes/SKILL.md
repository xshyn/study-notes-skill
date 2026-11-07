---
name: study-notes
description: Turn any learning resource — a URL, PDF, textbook chapter, lecture slides, Word doc, EPUB, YouTube video, notebook, or pasted text — into detailed study notes delivered as a polished Word document (.docx) AND a presentation deck (.pptx) for learning, review and exam preparation, with definitions, worked examples, formulas, comparison tables, diagrams, common mistakes, exam tips, cheat sheet, glossary, practice questions with an answer key, and flashcards. Use this skill whenever someone wants notes, a study guide, revision notes, lecture notes, a summary to study from, exam/test prep material, or slides to review a topic from a source they point to — even if they only say "make notes from this", "help me study this chapter", "I have an exam on this PDF", or "turn this article into slides", and even when they ask for only one of the two files.
---

# Study Notes

Produce study material a student can learn from, revise from, and prepare an exam from — faithful
to the source, complete, and well designed. One structured JSON file is the single source of truth;
bundled scripts render it into a Word document and a slide deck, so your effort goes into the
*content* and the layout is consistent every time.

```
source ──extract_source.py──▶ source.txt ──(you write)──▶ notes.json ──validate_notes.py──▶ fix gaps
                                                            │
                               build_docx.py ◀──────────────┴──────────────▶ build_pptx.py
                                     │                                            │
                                     └────────── render_preview.py (visual QA) ───┘
```

Scripts live in `scripts/` next to this file (call them by absolute path). They need Python 3.9+
with `python-docx` and `python-pptx`; everything else is optional with graceful fallbacks
(`pip install -r requirements.txt` for best results).

## Step 0 — Settle the brief (usually no questions needed)

Infer sensible defaults and start; only ask if something is genuinely ambiguous *and* expensive to
get wrong (e.g. a 600-page book with no indication of which chapters).

| decision | default |
|---|---|
| outputs | both `.docx` and `.pptx` (only one if the user asked for only one) |
| depth | `standard`; `exam` if they mention an exam/test/midterm/final; `deep` if they ask for comprehensive/everything; `summary` only if they ask for brief/short |
| language | the language the user wrote to you in, else the source's language |
| scope | the whole resource, or the pages/chapters/timestamps the user named |
| theme | `indigo` (any of indigo, teal, crimson, slate, forest, plum) |

Mention the defaults you chose in one line when you deliver.

## Step 1 — Get the full source text

```bash
python scripts/extract_source.py "<url-or-path>" -o work/source.txt --images work/figures
# PDFs: add --pages 40-75 to restrict to a chapter
```

Handles PDF, DOCX, PPTX, EPUB, HTML, Markdown/TXT, Jupyter notebooks, SRT/VTT subtitles, web pages
and YouTube transcripts. Output has `=== Page N ===` / `=== Slide N ===` markers — use them for
`source_ref`.

- **URL can't be fetched from the shell** (sandbox, paywall, JS-heavy page): fetch it with your own
  web tool (web fetch / browser), save the *complete* text to `work/source.md`, then continue. Make
  sure you have the whole page, not a summary — if your fetch tool summarises, request the full
  text in chunks.
- **Scanned PDF** (the script warns about very little text): OCR it (`ocrmypdf`), or read the
  page images directly if you can view images.
- **Read all of it.** Read `source.txt` completely before writing — in chunks for long sources
  (e.g. 300–500 lines at a time). Notes written from the first third of a document are the most
  common way this task goes wrong. While reading, jot the outline: headings → page ranges, every
  defined term, formula, example, figure, table and comparison.
- Look at extracted figures (`work/figures`) if you can view images; keep the ones that carry
  information (diagrams, plots, labelled schematics), skip logos and decoration.

## Step 2 — Write `notes.json`

**Read `references/pedagogy.md` (how to write notes that teach) and `references/notes_schema.md`
(the exact JSON format) before writing.** Key points:

- Mirror the source's structure: one top-level section per chapter/major heading, subsections for
  sub-headings, `source_ref` on each.
- Use the right block for each kind of content: `definition`, `example` (with steps and answer),
  `formula` (latex + plain `text` + variables), `table` for every comparison, `code`, `diagram`
  (Mermaid) for processes, `image` for figures worth keeping, and callouts — `key_point`,
  `warning` (common mistake), `tip` (exam tip), `mnemonic`, `intuition`, `note`.
- Be faithful: keep the source's terminology and notation; never invent facts, numbers, or
  citations. Your own added intuition or examples go in `intuition`/`note` callouts or clearly as
  extra examples.
- Back matter is not optional: overview, learning objectives, key takeaways, common mistakes, exam
  tips, cheat sheet, glossary, review questions (mixed types and difficulty, each tagged with its
  `section`, with full answers), flashcards.
- Say each thing once. A definition lives in its `definition` block; the body text builds on it
  rather than restating it, and recaps (takeaways, cheat sheet) compress rather than repeat.
- **Exam requests** ("midterm", "exam tomorrow", "what's likely to be asked"): add
  `exam_priorities` — a ranked list of topics (high / medium / low) with what to know for each.
  It opens the document and the deck. Give high-priority topics the most worked examples and
  questions; compress niche material (long lists of variants, history, minor algorithms) into one
  compact table instead of a section each. If the exam is very soon, prefer focus over volume.
- **Curate the deck while you write** — it is a review tool a student flips through, not the
  document re-flowed. Budget about **2–4 content slides per section and 20–35 slides in total**
  (up to ~45 for sources over 40 pages):
  - give every section `slide_bullets` — 3–6 crisp review bullets in your own words: what a
    student should recall on a quick flip-through;
  - mark with `"slide": true` only what earns a slide — the 1–2 key definitions, the one
    comparison table, the canonical worked example, the core formula or code, and the top
    mistake/tip. Prefer one strong visual per section (a `diagram`, a `chart` of how a quantity
    evolves, a table) over several text slides;
  - everything unmarked stays in the document and in the speaker notes, so nothing is lost.
  - each top-level section also costs ~2–3 fixed slides (divider, review/quiz, key-term cards
    when it has 3+ marked terms), so for sources with many short chapters, group related
    chapters into fewer top-level sections (keep the originals as subsections).
- For long sources write one file per chapter under `work/sections/` and list their paths in
  `sections` — smaller writes, easier fixes.

## Step 3 — Validate and close gaps

```bash
python scripts/validate_notes.py work/notes.json --source work/source.txt --depth standard
```

Fix every ERROR. Treat WARN lines as a to-do list: missing source headings, uncovered pages,
low key-term coverage, thin sections, too few questions — go back to the source and add the
missing material, then re-run. Don't pad to silence a warning; if a warning is genuinely not
applicable (e.g. a source with no formulas), leave it.

## Step 4 — Build the files

```bash
python scripts/build_docx.py work/notes.json -o "outputs/<Topic> - Study Notes.docx"
python scripts/build_pptx.py work/notes.json -o "outputs/<Topic> - Study Deck.pptx"
```

The document has a cover, overview/objectives, clickable contents, numbered sections with
"In brief" boxes and styled callouts, then takeaways, mistakes, tips, a cheat-sheet grid,
glossary, questions grouped by type with a separate answer key, and flashcards. **Everything in
the JSON goes into the document** — it is the complete reference.

The deck is the *review* companion. Per section it shows a divider, key-term cards (the
definitions you marked), the `slide_bullets` and marked content — short text and a small table
or code listing share a slide side by side; bigger examples, formulas, tables, figures and
diagrams get their own — then a "Remember" slide with the marked key points, mistakes and tips,
and a "Check yourself" quiz with answers in the speaker notes. After the sections come exam
priorities (up front), key takeaways, common mistakes, exam tips, the cheat sheet and a short
final quiz; the glossary stays in the document. Text is auto-sized and paginated (no overflow),
every slide carries speaker notes with the full explanation, and RTL languages are laid out
right-to-left automatically.

The builder prints the slide count and warns above 40. **Treat the budget as a hard target:** if
the deck is over it, un-mark secondary blocks and rebuild until it fits. If a section has
nothing but its divider, mark something or add `slide_bullets`. (Without any `slide_bullets` /
`"slide": true` markers the builder shows every block — a long deck — so always curate.)

## Step 5 — Look at the result

```bash
python scripts/render_preview.py "outputs/<Topic> - Study Notes.docx" "outputs/<Topic> - Study Deck.pptx" -o work/preview
```

This writes PDFs and contact-sheet PNGs (grids of page/slide thumbnails) and prints a text-overflow
check for the deck. If you can view images, open the contact sheets and check: nothing cut off,
figures visible, tables legible, no near-empty or overcrowded slides. Fix the JSON (e.g. add
`short` bullets for a wordy paragraph, split a huge table) and rebuild. If LibreOffice isn't
available, skip this step.

## Step 6 — Deliver

Give the user the two files, plus a 2–4 line summary: what was covered (sections / page range),
the counts (sections, practice questions, flashcards, glossary terms), and the defaults you chose.
Mention any part of the source you couldn't use (e.g. unreadable scanned pages, a paywalled
section). Don't paste the notes into the chat, and don't point the user to your intermediate
files (`source.txt`, `notes.json`) unless they asked for them.

## Requests that change the recipe

- **"Just slides" / "just a doc"** — still write the full `notes.json` (it's what makes either
  output good), build only the one asked for.
- **Several sources on one topic** — extract each, merge into one outline, cite each source in
  `source_ref` (e.g. "Lecture 3, slides 10–18; Textbook pp. 40–44").
- **An existing file to update** — load their previous `notes.json` if they have it; otherwise
  extract their old doc with `extract_source.py` and rebuild.
- **Other output formats** — Markdown/PDF: build the `.docx` and convert
  (`soffice --headless --convert-to pdf`), or write Markdown directly from the JSON.
- **Language not built in** (built-in labels: English, Persian, Arabic) — set `meta.labels` to
  translate headings like "Glossary", "Answer key" (see schema §2).
