# Lecture PPTX Skills

Claude Code skills (`.claude/skills/`) plus a python-pptx engine (`scripts/`) that turn a syllabus topic into an editable classroom lecture .pptx: teacher plan -> slide spec -> build -> QA.

- `lecture-workflow` is the entry point: "Generate the lecture PPT for <Unit/Topic>".
- `scripts/build_pptx.py` renders `slide_spec.json` (schema: `work/_TEMPLATE/slide_spec.example.json`).
- `scripts/qa_check.py` runs structural checks and renders slide PNGs via PowerPoint COM (Windows + MS PowerPoint required).
- House style: 16:9, nothing under 24pt, no filled blocks behind individual tokens. A senior reference deck is content-only, never a design reference.

Setup: copy `.claude/`, `scripts/`, `work/_TEMPLATE/` into a project; create `input/ work/ output/`; `pip install python-pptx Pillow matplotlib pywin32 pdfplumber "markitdown[pptx]"`.

Status: the engine's table layout sizing is still being tuned (a table + insight + example stack can overflow the slide).
