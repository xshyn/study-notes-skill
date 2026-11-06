# Writing notes that actually teach

Read this before writing the notes JSON. It is the difference between a summary and study material.

## The three jobs the notes must do

1. **Learn** — first contact with the material. Needs intuition before formalism, worked examples,
   and explanations of *why*, not just *what*.
2. **Review** — coming back a week later. Needs scannable structure, "In brief" summaries per
   section, key takeaways, and a cheat sheet.
3. **Exam prep** — the night before. Needs exam tips, common mistakes, formulas with variable
   meanings, and enough practice questions with answers to self-test.

A section that serves only one of these is unfinished.

## Fidelity to the source

- The source is the authority. Every claim, number, formula and definition should be traceable to
  it. Keep the source's terminology and notation — students will be examined in it.
- Where you add explanation the source lacks (an intuition, an analogy, a clarifying example),
  that's encouraged — mark it with an `intuition` or `note` callout so the learner knows it's an
  aid, not source content. Never invent data, dates, statistics, quotes, or citations.
- If the source is wrong or ambiguous, say so in a `note` callout rather than silently fixing it.
- Preserve the source's order unless it's genuinely confusing; learners cross-reference.
- Record where each section came from in `source_ref` (pages, slide numbers, timestamps, headings).

## Coverage: "detailed" means nothing important is dropped

- Every heading of the source should map to a section or subsection. Run the validator with
  `--source` and treat missing headings / uncovered pages as gaps to fill, not noise.
- Every defined term in the source → a `definition` block *and* a glossary entry.
- Every formula → a `formula` block with all variables explained and when it applies.
- Every worked example in the source → an `example` block with the steps kept. Add one of your own
  for concepts the source explains without an example.
- Every comparison (A vs B, pros/cons, types of X) → a `table`.
- Every process / algorithm → `numbered` steps, and a `diagram` if there are branches or a cycle.
- Every code listing that matters → a `code` block with an explanation of what to notice.
- Figures that carry information → re-use them (`image`) or redraw as a table/diagram.

Rough scale for `standard` depth: the notes text is typically 25–45% of the source's length for
textbook-dense material, and can exceed 100% for sparse sources (slides, short articles) where
you must add the explanation the lecturer gave verbally. `exam` and `deep` go further.

## Structure inside a section

A good default order (skip what doesn't apply):

1. `summary` — the gist in one or two sentences.
2. Motivation / intuition — why this exists, what problem it solves.
3. Definitions of new terms.
4. Core explanation (`text` / `bullets`) — mechanism, rules, properties.
5. Formulas with variables.
6. Worked example(s).
7. Comparison table if relevant.
8. `warning` (common mistake) and `tip` (exam tip) callouts.
9. `takeaways` for long sections.

Prefer several short blocks to one long paragraph. Keep bullets to one idea each (≤ ~30 words);
use children for detail.

## Practice questions (the biggest lever for exam prep)

- Mix types: recall (definitions, facts), conceptual (why/how), application (use it on a new case),
  problem (calculate/derive/trace), compare, and a few `mcq` / `true_false` with plausible
  distractors that target real misconceptions.
- Spread across difficulty (easy → hard) and across all sections; tag `section` on every question.
- Answers must be complete enough to grade yourself: the final value *and* the key step or reason.
  Add `explanation` when the answer alone wouldn't teach the reader why.
- Reuse the source's own end-of-chapter style if it has one.

Targets (validator checks these): summary ≥ 6, standard ≥ 12, exam ≥ 20, deep ≥ 25 questions;
at least 2–3 per major section.

## Flashcards, glossary, cheat sheet

- Flashcards: one atomic fact per card; front is a *question or cue*, not just a word
  ("What does a group of 4 cells eliminate?" beats "Group of 4").
- Glossary: every technical term, sorted automatically; definitions in one or two sentences.
- Cheat sheet: what you'd put on one allowed sheet of paper — formulas, rules, procedures,
  key numbers — grouped under headings, terse.

## Common mistakes and exam tips

- Common mistakes are the misconceptions a learner is *likely* to have — state the mistake and the
  correction. Pull them from the source's warnings, from known pitfalls of the topic, and from the
  traps your MCQ distractors target.
- Exam tips are actionable: what to write, what examiners look for, how to check an answer,
  which results to memorise vs derive.

## Language and tone

- Write the notes in the language the user asked for; default to the user's language if they wrote
  to you in it, otherwise the source's language. Keep technical terms in their standard form —
  when writing in a language other than the source's, give the original term in parentheses on
  first use (e.g. «ایمپلیکنت اول (prime implicant)»).
- Clear, direct, second-person where it helps ("Notice that…"). No filler like "In this section we
  will discuss…".

## Slides vs document

The same JSON feeds both, and the builders adapt it. Things that help the deck:
- Give long paragraphs a `short` list of slide bullets.
- Keep table cells short; put prose in `note`.
- One worked example per `example` block.
- Section `summary` doubles as the divider-slide text, so make it a real sentence.
