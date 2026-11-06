# Notes JSON schema

One JSON file drives both builders (`build_docx.py`, `build_pptx.py`). Write it once, render twice.
Inline text anywhere supports light markdown: `**bold**`, `*italic*`, `` `code` ``, `==highlight==`
(don't nest them).

## Contents
1. Top level
2. `meta`
3. Sections
4. Block types (with examples)
5. Back matter: takeaways, mistakes, tips, cheat sheet, glossary, questions, flashcards
6. Long sources: per-section files
7. Minimal complete example

---

## 1. Top level

```json
{
  "meta": { ... },
  "overview": "2–4 sentence big-picture paragraph (string or list of strings)",
  "learning_objectives": ["Action-verb outcomes: Explain…, Derive…, Compare…"],
  "prerequisites": ["What the learner should already know"],
  "sections": [ ... ],
  "key_takeaways": ["…"],
  "common_mistakes": [{"mistake": "…", "correction": "…"}],
  "exam_tips": ["…"],
  "cheat_sheet": [{"heading": "…", "items": ["…"]}],
  "glossary": [{"term": "…", "definition": "…"}],
  "review_questions": [ ... ],
  "flashcards": [{"front": "…", "back": "…"}],
  "further_reading": ["…"]
}
```

## 2. `meta`

| key | meaning |
|---|---|
| `title` | **required**. Topic name, not the file name |
| `subtitle` | course / chapter / lecture |
| `source` | human-readable citation: "Author, *Title*, ch. 3 (pp. 70–98)" or the URL |
| `author`, `course`, `level`, `date` | shown on the cover |
| `language` | ISO code (`en`, `fa`, `ar`, `de`…). RTL languages (fa, ar, he, ur, ps, ckb…) flip layout automatically |
| `depth` | `summary` · `standard` · `exam` · `deep` — used by the validator's targets |
| `theme` | `indigo` (default) · `teal` · `crimson` · `slate` · `forest` · `plum` |
| `colors` | optional overrides, e.g. `{"primary": "#123456", "accent": "#F5A623"}` |
| `font_body`, `font_heading` | override fonts (default Calibri; Vazirmatn for RTL) |
| `labels` | override any UI label (e.g. for German: `{"glossary": "Glossar", "answers": "Lösungen"}`). English, Persian and Arabic labels are built in |
| `paper` | `letter` (default) or `a4` |
| `short_title` | footer text on slides if the title is long |
| `section_page_breaks` | `false` to run sections together in the doc |
| `section_quizzes` | `false` to omit the per-section "Check yourself" slides |

## 3. Sections

```json
{
  "title": "Building and reading a K-map",
  "source_ref": "pp. 73–85",
  "summary": "One or two sentences: the gist of this section.",
  "blocks": [ ... ],
  "takeaways": ["optional per-section recap bullets"],
  "subsections": [ { "title": "…", "summary": "…", "blocks": [ ... ] } ]
}
```

- `source_ref` — pages / slides / timestamps / headings of the source this section covers.
  The validator uses page ranges here to report pages you never covered.
- `summary` — becomes the "In brief" box in the doc and the subtitle of the slide divider.
- Subsections nest to any depth (rendered as 2.1, 2.1.1 …).

## 4. Block types

Every block has a `type`. Unknown types render as text.

### text
```json
{"type": "text", "text": "Full explanatory paragraph…", "short": ["Optional slide bullet 1", "bullet 2"]}
```
On slides, a paragraph is split into one bullet per sentence unless you give `short` (recommended
for paragraphs longer than ~3 sentences). The full text always goes into the speaker notes.

### bullets / numbered (alias `steps`)
```json
{"type": "bullets", "title": "optional lead-in", "items": ["plain item", {"text": "parent", "children": ["child"]}]}
```

### definition
```json
{"type": "definition", "term": "Prime implicant", "aka": "PI", "definition": "…",
 "example": "optional short example", "note": "optional caveat"}
```
Consecutive definitions are grouped into a card grid on slides (up to 4 per slide).

### Callouts: `key_point`, `warning` (common mistake), `tip` (exam tip), `mnemonic`, `note`, `intuition`
```json
{"type": "warning", "text": "…", "title": "optional custom label", "items": ["optional bullets"]}
```
Consecutive callouts share a slide (up to 3).

### example (worked example)
```json
{"type": "example", "title": "Simplify F = Σm(0,2,5,7)",
 "problem": "Statement of the problem",
 "steps": ["Step 1 with reasoning", "Step 2 …"],
 "code": "optional code shown with the problem",
 "solution": "optional prose solution",
 "answer": "Final answer, stated plainly"}
```
Slides put the problem on the left and the numbered solution on the right.

### formula
```json
{"type": "formula", "name": "Ohm's law", "latex": "V = I R", "text": "V = I·R",
 "variables": [{"symbol": "V", "meaning": "voltage (volts)"}],
 "explanation": "What it means and when it applies."}
```
`latex` is rendered to an image with matplotlib mathtext (a large LaTeX subset; common macros
like `\iff`, `\text{}`, `\dfrac` are auto-translated). Always add `text` as a plain fallback.

### code
```json
{"type": "code", "language": "python", "title": "Binary search", "code": "def …",
 "explanation": "What to notice in this code."}
```
Long listings are split across slides automatically.

### table
```json
{"type": "table", "title": "TCP vs UDP", "headers": ["Aspect", "TCP", "UDP"],
 "rows": [["Reliability", "Guaranteed", "Best-effort"]], "note": "optional footnote"}
```
Comparison tables are one of the highest-value study aids — use them whenever the source compares
two or more things. Tall tables paginate on slides.

### image / figure
```json
{"type": "image", "path": "figures/page012_img1.png", "caption": "Fig 3.4 — 4-variable K-map", "width_in": 5}
```
Paths are relative to the notes JSON. `extract_source.py --images DIR` saves figures from PDFs,
DOCX and PPTX for reuse.

### diagram
```json
{"type": "diagram", "mermaid": "flowchart LR\n A-->B", "caption": "…", "ascii": "optional fallback"}
```
Rendered with mermaid-cli (`mmdc`) when installed; otherwise the `ascii` fallback (or the Mermaid
source) is shown as a code block. Use for processes, hierarchies, state machines, timelines.

### quote
```json
{"type": "quote", "text": "…", "by": "Author"}
```

## 5. Back matter

```json
"review_questions": [
  {"question": "…", "answer": "…", "explanation": "optional why",
   "type": "recall | conceptual | application | problem | compare | analysis | mcq | true_false",
   "difficulty": "easy | medium | hard",
   "options": ["only for mcq"],
   "section": "2"}
]
```
- Tag every question with the top-level `section` number it tests: the deck shows up to 3 tagged
  questions as a "Check yourself" quiz after that section; the rest go in the final quiz.
- The doc prints questions grouped by type and sorted by difficulty, with a separate answer key.

`common_mistakes` items may be strings or `{"mistake", "correction"}` objects (preferred).
`cheat_sheet` items may be strings or `{"heading", "items"}` groups (preferred — renders as a grid).

## 6. Long sources: per-section files

Any entry in `sections` may be a path to a JSON file holding one section object (or a list of
them), relative to the main notes file:

```json
"sections": ["sections/01-intro.json", "sections/02-kmaps.json"]
```
Write one file per chapter as you read; keeps each write small and lets you revise one chapter
without rewriting everything.

## 7. Minimal complete example

```json
{
  "meta": {"title": "TCP Congestion Control", "source": "Kurose & Ross §3.7", "language": "en", "depth": "standard"},
  "overview": "TCP adjusts its sending rate to avoid overwhelming the network…",
  "learning_objectives": ["Explain slow start and congestion avoidance", "Trace cwnd over time"],
  "sections": [{
    "title": "Slow start", "source_ref": "pp. 270–272", "summary": "cwnd doubles each RTT until loss or ssthresh.",
    "blocks": [
      {"type": "definition", "term": "cwnd", "definition": "Congestion window: max unacknowledged bytes in flight."},
      {"type": "text", "text": "…"},
      {"type": "example", "title": "cwnd growth", "steps": ["RTT 1: 1 MSS", "RTT 2: 2 MSS", "RTT 3: 4 MSS"], "answer": "Exponential growth"},
      {"type": "warning", "text": "Slow start is not slow — growth is exponential."}
    ]}],
  "key_takeaways": ["…"],
  "glossary": [{"term": "MSS", "definition": "Maximum segment size"}],
  "review_questions": [{"question": "Why is it called slow start?", "answer": "…", "type": "conceptual", "difficulty": "medium", "section": "1"}],
  "flashcards": [{"front": "ssthresh", "back": "Threshold where slow start ends"}]
}
```
