---
name: lecture-pptx-structure
description: Turn a lesson plan into a slide-by-slide structure for a classroom lecture deck. Adapted from Gabberflast/academic-pptx-skill (research/conference presentation structure) into classroom-teaching structure. Run after lecture-teacher produces lesson_plan.md, before lecture-pptx-build generates the actual .pptx. Produces slide_spec.json.
disable-model-invocation: false
---

# Lecture PPTX Structure Skill

Converts `work/<unit-slug>/lesson_plan.md` into `work/<unit-slug>/slide_spec.json`,
the exact JSON contract `scripts/build_pptx.py` consumes. Schema and every layout
type are documented in `work/_TEMPLATE/slide_spec.example.json` — read it before
writing a spec.

This is a classroom deck, not a conference talk: optimize for a teacher explaining
live in a room, not for a reader skimming alone. Slides are prompts for speech, not
the speech itself.

## Step 1 — Decide the flow per sub-topic, not a fixed template

For each sub-topic in the lesson plan, choose which stages actually apply from:

```
Context → Prerequisite refresher → Concept → Explanation → Diagram/Visual →
Example → Worked Example → Common Mistake → Real-world Application → Recap
```

A trivial sub-topic might only need Concept → Example. A hard one might need all
of them plus two worked examples. **Do not force every sub-topic through the same
number of stages or the same slide count.** The lesson plan's per-sub-topic
"Depth" and "Diagram needed?" fields already tell you how much a sub-topic needs —
honor them.

## Step 2 — Map content to layouts

Available layouts (see template for exact JSON fields per layout):

| Layout | Use for |
|---|---|
| `title` | Unit cover slide |
| `section` | Sub-topic / chapter transition |
| `content` | Concept explanation, bulleted |
| `two_column` | Prerequisite-vs-new, before/after, theory-vs-practice |
| `comparison` | Head-to-head comparison of two approaches/options |
| `diagram` | Process flow, architecture, sequence — native editable boxes+arrows |
| `table` | Structured comparison data, truth tables, complexity tables |
| `image` | When a photo/screenshot/reference image genuinely helps (from `input/`) |
| `formula` | A definition or formula that needs to stand alone, explained |
| `code` | Worked code example |
| `mistake` | Common student mistake + the fix |
| `token_categories` | Classify a small fixed set of categories (e.g. C tokens: Keywords/Identifiers/Constants/Strings/Operators/Special Symbols) as clean bordered pills + a definition callout |
| `token_breakdown` | Break one worked example into its category-colored parts (e.g. `int age = 20;` → Keyword/Identifier/Operator/Constant/Special Symbol cards) |
| `recap` | End-of-lecture recap questions |

## Step 3 — Apply the house design system (never the senior reference deck)

**The senior/faculty reference deck in `input/` must have zero influence on
design.** Not its colors, fonts, aspect ratio, layouts, slide structure, or
information density. It was provided as content material only (see
`lecture-teacher/SKILL.md` §Inputs) — do not open it, or `reference_style/`, at
this stage. `reference_style/` is deprecated; read its `NOTE.md` if you're unsure
why.

Instead, every unit uses this project's own modern academic **house style**,
defined once in `scripts/build_pptx.py`'s `DEFAULT_THEME` and applied consistently
across all units so the course looks like one coherent product:

- 16:9 widescreen (13.333in × 7.5in) — matches modern classroom projectors
- Segoe UI throughout, **classroom-projector font floor — nothing on any slide
  renders below 24pt, ever** (a real correction from a user who couldn't read a
  generated slide from the back of the room; this is a standing rule, not a
  per-slide judgment call):
  - Slide title: 34-40pt (`size_title`/`size_section`)
  - Section headings: 28-32pt (`size_heading`)
  - Main explanatory text: 26-30pt (`size_body`)
  - Labels / captions / footers: 24-26pt minimum (`size_small`/`size_caption`)
    — de-emphasize a footer or caption with color/weight, never by shrinking it
    below 24pt
  - Code examples: 28-32pt (`size_code`)
  - **If content doesn't fit at these sizes, split into two slides. Never drop
    a size below its floor to make content fit** — this is the fallback the
    house style always takes.
- Deep teal-navy primary (`0B3D5C`) + amber accent (`E08E1D`) + burnt-orange
  emphasis (`C1440E`) for mistakes/warnings — high-contrast, colorblind-safe,
  professional, not decorative. A fixed six-color palette (`token_colors` in
  `DEFAULT_THEME`) exists for token/category-style content (keywords,
  identifiers, operators, etc.) — reuse it whenever a slide classifies items
  into a small fixed set of categories, so color-coding stays consistent
  across the course rather than reinvented per slide.
- Generous whitespace with content sized to actually fill the slide (see Hard
  Rules below) rather than either cramming or leaving dead space
- **No filled color blocks behind individual words/tokens/values** — use clean
  typography, light backgrounds, subtle borders, or a thin colored accent bar
  instead. Reserve solid color fills for large intentional surfaces (a title
  slide's full background, a section divider) — never as a box wrapped tightly
  around one short label, which reads as cluttered rather than emphasized.
- **University logo placeholder**: any slide can set `"show_logo": true` (the
  `token_categories`/`token_breakdown` layouts default to it) to reserve the
  top-right corner for a logo via `Deck.add_logo_placeholder()` — currently a
  clearly-marked placeholder box since no real logo file has been supplied yet.
  When the user provides an actual logo image, swap the placeholder rectangle
  for `slide.shapes.add_picture(...)` at the same coordinates it returns.

You may adjust `meta.theme` per-unit only for genuine pedagogical reasons (e.g. a
unit that's unusually diagram-heavy), never to imitate an input document. If you
want to evolve the house style itself (new accent color, different type scale),
change the shared `DEFAULT_THEME` in `build_pptx.py` so it stays consistent across
every unit, rather than a one-off override in a single `slide_spec.json`.

**Fill the slide.** A layout with only 2-3 short items and a lot of blank lower
half is a bug, not a minimalist choice — either the content needs more (a worked
example, a table, a diagram) or the element sizing needs to grow to use the space
well. `scripts/build_pptx.py`'s existing layouts (`diagram` sublabels/intro,
`table` insight/example callouts, `code` output preview, `recap` closing banner)
exist specifically to fill a slide meaningfully instead of leaving it sparse —
use them.

## Hard rules (content density & quality)

1. **No walls of text.** Max ~6 bullets per `content` slide, each bullet a short
   phrase, not a sentence a student has to read instead of listening.
2. **Every slide earns its place.** A slide with only 2-3 short bullets and nothing
   else should be merged into a neighboring slide or given a diagram/table/code
   example instead.
3. **A worked example within 1-2 slides of every new definition or formula.**
   Never leave a definition standalone.
4. **Diagrams over prose** whenever the lesson plan flagged "Diagram needed: yes" —
   don't describe a process in bullets when a `diagram` layout shows it directly.
5. **Layout variety.** Don't use the same layout 3+ times in a row within one
   sub-topic; alternate content/diagram/table/code/comparison as the material
   allows — but never force a layout the content doesn't warrant.
6. **No decorative filler.** No slide exists just to look impressive. No stock-photo
   feel, no generic "Thank You" filler slides mid-deck.
7. **Speaker notes carry the explanation.** Put the professor's actual spoken
   explanation, analogy, and transition into each slide's `notes` field — the
   visible slide stays terse; the notes are what makes it teachable live.
8. **Recap slide(s) at the end of the unit**, built from the lesson plan's recap
   questions.
9. **Slide count is a consequence, not a target.** Do not pad to hit a round
   number, and do not compress hard material to save slides.

## Step 4 — Self-check before handing off

- [ ] Every lesson-plan sub-topic is represented
- [ ] Every worked example from the lesson plan appears as a `code`/`content`/
      `formula` slide with the actual worked steps, not a placeholder
- [ ] Every flagged "common mistake" has a `mistake` slide
- [ ] Every flagged diagram has a `diagram` slide
- [ ] No 3+ consecutive slides use the same layout
- [ ] JSON validates against the schema (run
      `.venv\Scripts\python -c "import json; json.load(open('work/<unit-slug>/slide_spec.json'))"`)

## Handoff

Once `slide_spec.json` is written and validated, hand off to `lecture-pptx-build`
to render, render-QA, and fix.
