---
name: lecture-workflow
description: End-to-end orchestrator for generating a classroom lecture PowerPoint from a syllabus unit/topic. Trigger whenever the user says "generate the lecture PPT for Unit X / Topic X", "make slides for [topic]", or similar — this is the main user-facing entry point for the whole pipeline (Teacher → Structure → Build → QA).
argument-hint: "Unit/topic name or number, e.g. 'Unit 1 - Introduction to C'"
disable-model-invocation: false
---

# Lecture Workflow (Orchestrator)

This is the reusable pipeline: syllabus/topic in → reviewed, editable `.pptx` out.
Run every step below in order. Do not skip QA. Do not generate slides before the
teaching plan is approved conceptually (you don't need to stop and ask the user at
every stage, but the plan must exist and be self-consistent before slides start).

## 0. Setup (once per shell session)

```bash
source ./env.sh
```
(or `. .\env.ps1` in PowerShell). This redirects pip/matplotlib/npm/temp caches to
`.cache/` on Drive D — never skip this before running any script in this project.

If `.venv` doesn't exist yet or a required package is missing:
```bash
.venv\Scripts\python -m pip install --cache-dir .cache\pip -q python-pptx Pillow matplotlib pywin32 pdfplumber "markitdown[pptx]"
```

## 1. Locate the unit/topic

- Read the syllabus in `input/` (PDF — use `pdfplumber` via a short Python snippet,
  or the `pdf` skill, to extract the relevant unit's text) and find the exact scope
  for the requested unit/topic.
- Check `input/` for any topic notes, resource files, or images related to it.
- Determine a `unit-slug` (e.g. `unit-1-intro-to-c`) — all intermediate work for
  this run goes in `work/<unit-slug>/`.

## 2. Senior reference deck — content only, never design

**Any senior/faculty reference `.pptx` in `input/` is a secondary CONTENT
reference only** — what was previously taught, terminology, examples, topic
coverage. It is explicitly **not** a design, layout, density, or quality
reference. Never copy or imitate its colors, fonts, layouts, slide structure, or
information density, and never let it cap the quality or creativity of the new
deck.

If `reference_content/<slug>_content.md` doesn't exist yet for that deck, generate
it:
```bash
.venv\Scripts\python scripts\extract_ref_content.py "input\<reference>.pptx"
```
Read it once, during Step 3 (Teacher Skill), purely to sanity-check terminology
and coverage — then verify and improve on it, don't reuse it verbatim. Do not
generate or read `reference_style/` — that folder is deprecated (see its
`NOTE.md`). Visual design for every unit comes from the fixed house style in
`scripts/build_pptx.py`'s `DEFAULT_THEME`, not from any input deck.

## 3. Teacher Skill → lesson plan

Invoke the `lecture-teacher` skill for this unit/topic. It reads the syllabus scope
+ notes + reference style pacing, and writes `work/<unit-slug>/lesson_plan.md`.

## 4. Academic PPTX Structure Skill → slide spec

Invoke the `lecture-pptx-structure` skill. It reads `lesson_plan.md` and applies
the project's fixed house design style (never the senior reference deck) to write
`work/<unit-slug>/slide_spec.json` (schema:
`work/_TEMPLATE/slide_spec.example.json`).

## 5. PowerPoint Generation Skill → build + QA + fix

Invoke the `lecture-pptx-build` skill. It runs:
```bash
.venv\Scripts\python scripts\build_pptx.py work\<unit-slug>\slide_spec.json output\<unit-slug>.pptx
.venv\Scripts\python scripts\qa_check.py output\<unit-slug>.pptx work\<unit-slug>\qa
```
then visually inspects the rendered slide images and fixes issues by editing
`slide_spec.json` and rebuilding, looping until the QA checklist passes (max 3
rounds).

## 6. Final coverage review

Before reporting done, re-read `lesson_plan.md` against `slide_spec.json` /
`output/<unit-slug>.pptx` and confirm:
- [ ] Every syllabus topic for this unit is covered
- [ ] Every lesson-plan sub-topic appears, in the planned teaching order
- [ ] Every prerequisite refresher, worked example, and common-mistake slide from
      the plan made it into the deck
- [ ] No text overflow / overlap / cut-off content (from Step 5's QA)
- [ ] Formulas render correctly (image-based — note this to the user)
- [ ] Visual design is the project's own modern house style — not imitating the
      senior reference deck's colors/layout/density
- [ ] It would actually work as a live teaching aid, not just a readable document

If anything is missing, go back to the relevant step (don't patch the .pptx by
hand) and rebuild.

## 7. Report to the user

State: output path (`output/<unit-slug>.pptx`), slide count, what's covered, any
known limitations, and where intermediate artifacts live
(`work/<unit-slug>/lesson_plan.md`, `slide_spec.json`, `qa/`) in case the user wants
to inspect or hand-tune before the next regeneration.

## Usage (for the user)

Once materials are in `input/`, the whole pipeline runs from one instruction:

> "Generate the lecture PPT for Unit 1 - Introduction to C"

New units just need their syllabus/notes present in `input/` — no reconfiguration.
