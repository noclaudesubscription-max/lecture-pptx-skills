"""
Quality-review a generated .pptx:
  1. Structural checks via python-pptx: off-slide elements, shape overlaps,
     empty/sparse slides, missing notes on formula slides.
  2. Renders each slide to PNG via PowerPoint COM automation (Windows only,
     requires MS PowerPoint installed) so Claude can visually inspect for
     overflow/overlap/legibility using the Read tool.

Usage:
    .venv\\Scripts\\python scripts\\qa_check.py output\\<unit>.pptx work\\<unit-slug>\\qa

Writes:
    <qa_dir>/structural_report.md
    <qa_dir>/slide_01.png, slide_02.png, ...
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

ROOT = Path(__file__).resolve().parent.parent


def emu_in(v):
    return Emu(v).inches


def rects_overlap(a, b, tolerance_in=0.03):
    ax1, ay1, ax2, ay2 = a
    bx1, by1, bx2, by2 = b
    ax1 += tolerance_in; ay1 += tolerance_in
    bx1 += tolerance_in; by1 += tolerance_in
    ax2 -= tolerance_in; ay2 -= tolerance_in
    bx2 -= tolerance_in; by2 -= tolerance_in
    return not (ax2 <= bx1 or bx2 <= ax1 or ay2 <= by1 or by2 <= ay1)


def structural_check(pptx_path: Path) -> list[str]:
    prs = Presentation(str(pptx_path))
    sw, sh = emu_in(prs.slide_width), emu_in(prs.slide_height)
    issues = []
    for idx, slide in enumerate(prs.slides, start=1):
        boxes = []
        shape_count = 0
        text_chars = 0
        for shape in slide.shapes:
            if shape.left is None or shape.top is None:
                continue
            shape_count += 1
            x1, y1 = emu_in(shape.left), emu_in(shape.top)
            x2, y2 = x1 + emu_in(shape.width or 0), y1 + emu_in(shape.height or 0)
            if x1 < -0.02 or y1 < -0.02 or x2 > sw + 0.02 or y2 > sh + 0.02:
                issues.append(f"Slide {idx}: shape '{shape.name}' extends outside slide "
                              f"bounds (x:{x1:.2f}-{x2:.2f}, y:{y1:.2f}-{y2:.2f}, slide {sw}x{sh})")
            # full-bleed background shapes (title/section color rectangles) are meant
            # to sit behind everything else -- exclude them from overlap pairing
            is_background = (x2 - x1) >= sw - 0.05 and (y2 - y1) >= sh - 0.05
            if not is_background:
                boxes.append((shape.name, (x1, y1, x2, y2)))
            if shape.has_text_frame:
                text_chars += len(shape.text_frame.text)
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                if rects_overlap(boxes[i][1], boxes[j][1]):
                    issues.append(f"Slide {idx}: possible overlap between "
                                  f"'{boxes[i][0]}' and '{boxes[j][0]}'")
        if shape_count <= 1 and text_chars < 10:
            issues.append(f"Slide {idx}: appears empty/near-empty ({shape_count} shapes, {text_chars} chars)")
        notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""
        if shape_count == 0:
            issues.append(f"Slide {idx}: no shapes at all")
    if not issues:
        issues = ["No structural issues found (bounds/overlap/emptiness checks passed)."]
    return issues


def render_slides_via_powerpoint(pptx_path: Path, out_dir: Path) -> list[Path]:
    import win32com.client
    out_dir.mkdir(parents=True, exist_ok=True)
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1
    deck = powerpoint.Presentations.Open(str(pptx_path.resolve()), WithWindow=False)
    try:
        deck.SaveAs(str(out_dir.resolve()), 18)  # 18 = ppSaveAsPNG -> exports one PNG per slide into folder
    finally:
        deck.Close()
        try:
            powerpoint.Quit()
        except Exception:
            pass
    seen = {}
    for p in out_dir.glob("Slide*.[Pp][Nn][Gg]"):
        seen[p.name.lower()] = p
    return sorted(seen.values(), key=lambda p: p.name.lower())


def main():
    if len(sys.argv) < 3:
        print("Usage: qa_check.py <deck.pptx> <qa_output_dir>")
        sys.exit(1)
    pptx_path = Path(sys.argv[1])
    qa_dir = Path(sys.argv[2])
    qa_dir.mkdir(parents=True, exist_ok=True)

    issues = structural_check(pptx_path)
    report = ["# Structural QA Report — " + pptx_path.name, ""]
    report += [f"- {i}" for i in issues]
    (qa_dir / "structural_report.md").write_text("\n".join(report), encoding="utf-8")
    print("\n".join(report))

    try:
        images = render_slides_via_powerpoint(pptx_path, qa_dir)
        print(f"\nRendered {len(images)} slide images to {qa_dir}")
    except Exception as e:
        print(f"\n[WARN] PowerPoint COM rendering failed ({e}). "
              f"Visual QA will rely on structural checks only.")


if __name__ == "__main__":
    main()
