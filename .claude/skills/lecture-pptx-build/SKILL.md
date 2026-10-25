---
name: lecture-pptx-build
description: Render a slide_spec.json into a fully editable .pptx, then render-QA it and fix issues. Adapted from Noi1r/powerpoint-skill (which uses pptxgenjs/LaTeX/LibreOffice on macOS/Linux) into a Windows-native, python-pptx + PowerPoint-COM pipeline, since this machine has no LaTeX/Graphviz/LibreOffice toolchain but does have MS PowerPoint installed. Run after lecture-pptx-structure produces slide_spec.json.
disable-model-invocation: false
---

# Lecture PPTX Build Skill

Turns `work/<unit-slug>/slide_spec.json` into `output/<unit-slug>.pptx`, using the
project's own render engine (`scripts/build_pptx.py`) rather than pptxgenjs/LaTeX —
this machine has no Node/LaTeX/Graphviz/LibreOffice toolchain, but does have MS
PowerPoint, which we use for QA rendering via COM automation instead of `soffice`.

Always `source ./env.sh` (or `. .\env.ps1` in PowerShell) first, in every shell that
runs any of these scripts, so pip/matplotlib/npm/temp caches stay on Drive D.

## Step 1 — Build

```bash
.venv\Scripts\python scripts\build_pptx.py work\<unit-slug>\slide_spec.json output\<unit-slug>.pptx
```

This is deterministic — all text, shapes, tables, and diagrams it produces are
native PowerPoint objects (fully editable). The one exception: formulas are
rendered as a high-DPI transparent PNG via matplotlib mathtext (no LaTeX install
needed on this machine), because python-pptx has no OMML/native-math writer. The
plain-text formula is always also stored in that slide's speaker notes so it can be
retyped/edited by hand if needed — mention this limitation in your final report to
the user.

## Step 2 — Render-QA

```bash
.venv\Scripts\python scripts\qa_check.py output\<unit-slug>.pptx work\<unit-slug>\qa
```

This does two things:
1. **Structural checks** (`work\<unit-slug>\qa\structural_report.md`): shapes
   extending outside slide bounds, shape-to-shape overlaps, near-empty slides.
2. **Visual render**: opens the deck in PowerPoint via COM automation and exports
   one PNG per slide to `work\<unit-slug>\qa\SlideN.PNG` — read these with the
   `Read` tool to visually inspect. **Never read more than ~6-7 slide images in the
   main conversation** — for decks bigger than that, dispatch parallel `Explore` or
   general-purpose subagents, each assigned 5 slides, and have them return a
   text-only issues report (this mirrors the source skill's context-budget rule).

## Step 3 — QA checklist (apply per slide image)

- [ ] No text/shape clipped at any edge
- [ ] No overlapping elements (besides intentional full-bleed title/section
      background rectangles)
- [ ] No word wraps mid-word into an ugly orphan (e.g. a lone trailing "s" on
      its own line) — widen the box, shorten the label, or use a grid instead
      of a single cramped row
- [ ] **Nothing on the slide is smaller than 24pt** — this is an absolute floor,
      not a guideline (see `lecture-pptx-structure/SKILL.md` §3 for the full
      title/heading/body/label/code bands). If content doesn't fit at these
      sizes, that's a signal to split the slide, not shrink text below 24pt.
- [ ] Diagrams: boxes/arrows don't overlap, labels fit inside their box
- [ ] Formula image is crisp, not blurry/pixelated, correctly centered
- [ ] Tables fit the slide width, header row distinguishable from data rows
- [ ] Content fills the slide reasonably — no slide mostly blank, none overflowing
- [ ] No filled color blocks wrapped tightly around individual words/tokens —
      use light backgrounds/borders/accent bars instead (house style, not just
      for token-teaching slides)
- [ ] Visual style consistent with the project's own house style (16:9,
      teal/amber, Segoe UI — see `lecture-pptx-structure/SKILL.md` §3), never the
      senior reference deck's look

## Step 4 — Fix loop

If `structural_report.md` or visual inspection finds issues, the fix is almost
always in `work\<unit-slug>\slide_spec.json` (adjust bullet count, split a
crowded slide into two, shorten a table, change a layout) — edit the spec, not the
.pptx directly, then re-run Step 1 and Step 2. Re-editing the generated .pptx by
hand is a dead end since the next rebuild would overwrite it.

If a genuinely new visual need appears (a layout the spec schema can't express),
extend `scripts\build_pptx.py` with a new `layout_*` method and register it in
`Deck.build()`'s `mapping` dict, following the existing layout methods' pattern
(title bar → content region → optional footer/notes). Keep new layouts consistent
with the existing theme dict rather than hardcoding colors/sizes.

Repeat build → QA → fix until the checklist passes clean, capped at 3 rounds
(matching the source skill's QA loop budget) — if still failing after 3 rounds,
report the remaining issues to the user rather than looping indefinitely.

## Step 5 — Report

State the output path, slide count, file size, and a one-line summary of anything
fixed or any known limitation (e.g. formulas are images, not native OMML).
