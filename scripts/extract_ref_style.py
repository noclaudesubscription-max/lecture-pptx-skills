"""
Analyze a reference .pptx (senior-faculty deck) and dump its visual style
(slide size, fonts, sizes, colors, layout density, placeholder geometry)
to reference_style/style_profile.json + a human-readable .md summary.

Usage:
    .venv\\Scripts\\python scripts\\extract_ref_style.py input\\"Unit 1 - Intro to C (1).pptx"
"""
import json
import sys
from collections import Counter
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

ROOT = Path(__file__).resolve().parent.parent


def emu_to_in(v):
    return round(Emu(v).inches, 3) if v is not None else None


def rgb_of(color_format):
    try:
        if color_format.type is not None:
            return str(color_format.rgb)
    except Exception:
        return None
    return None


def analyze(pptx_path: Path) -> dict:
    prs = Presentation(str(pptx_path))
    profile = {
        "source_file": pptx_path.name,
        "slide_width_in": emu_to_in(prs.slide_width),
        "slide_height_in": emu_to_in(prs.slide_height),
        "aspect_ratio": round(prs.slide_width / prs.slide_height, 4),
        "slide_count": len(prs.slides.__iter__.__self__._sldIdLst),
        "fonts": Counter(),
        "font_sizes_pt": Counter(),
        "title_font_sizes_pt": Counter(),
        "body_font_sizes_pt": Counter(),
        "colors": Counter(),
        "shape_types": Counter(),
        "avg_shapes_per_slide": 0,
        "slides_with_images": 0,
        "slides_with_tables": 0,
        "slides_with_charts": 0,
        "layout_names_used": Counter(),
    }

    total_shapes = 0
    n_slides = 0
    for slide in prs.slides:
        n_slides += 1
        try:
            profile["layout_names_used"][slide.slide_layout.name] += 1
        except Exception:
            pass
        has_image = has_table = has_chart = False
        for shape in slide.shapes:
            total_shapes += 1
            profile["shape_types"][str(shape.shape_type)] += 1
            if shape.shape_type == 13:  # PICTURE
                has_image = True
            if shape.has_table:
                has_table = True
            if shape.has_chart:
                has_chart = True
            if not shape.has_text_frame:
                continue
            is_title = shape == slide.shapes.title
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.name:
                        profile["fonts"][run.font.name] += 1
                    if run.font.size:
                        pt = int(run.font.size.pt)
                        profile["font_sizes_pt"][pt] += 1
                        if is_title:
                            profile["title_font_sizes_pt"][pt] += 1
                        else:
                            profile["body_font_sizes_pt"][pt] += 1
                    c = rgb_of(run.font.color)
                    if c:
                        profile["colors"][c] += 1
        if has_image:
            profile["slides_with_images"] += 1
        if has_table:
            profile["slides_with_tables"] += 1
        if has_chart:
            profile["slides_with_charts"] += 1

    profile["slide_count"] = n_slides
    profile["avg_shapes_per_slide"] = round(total_shapes / max(n_slides, 1), 2)
    for key in ("fonts", "font_sizes_pt", "title_font_sizes_pt", "body_font_sizes_pt",
                "colors", "shape_types", "layout_names_used"):
        profile[key] = dict(profile[key].most_common(15))
    return profile


def to_markdown(profile: dict) -> str:
    lines = [f"# Reference Style Profile — {profile['source_file']}", ""]
    lines.append(f"- Slide size: {profile['slide_width_in']}in x {profile['slide_height_in']}in "
                  f"(aspect {profile['aspect_ratio']})")
    lines.append(f"- Slide count: {profile['slide_count']}")
    lines.append(f"- Avg shapes/slide: {profile['avg_shapes_per_slide']}")
    lines.append(f"- Slides with images: {profile['slides_with_images']}, "
                 f"tables: {profile['slides_with_tables']}, charts: {profile['slides_with_charts']}")
    lines.append("")
    lines.append("## Fonts (most used)")
    for k, v in profile["fonts"].items():
        lines.append(f"- {k}: {v} runs")
    lines.append("")
    lines.append("## Title font sizes (pt)")
    for k, v in profile["title_font_sizes_pt"].items():
        lines.append(f"- {k}pt: {v} runs")
    lines.append("")
    lines.append("## Body font sizes (pt)")
    for k, v in profile["body_font_sizes_pt"].items():
        lines.append(f"- {k}pt: {v} runs")
    lines.append("")
    lines.append("## Colors (hex, most used)")
    for k, v in profile["colors"].items():
        lines.append(f"- #{k}: {v} runs")
    lines.append("")
    lines.append("## Layouts used")
    for k, v in profile["layout_names_used"].items():
        lines.append(f"- {k}: {v} slides")
    return "\n".join(lines)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: extract_ref_style.py <path-to-reference.pptx>")
        sys.exit(1)
    src = Path(sys.argv[1])
    profile = analyze(src)
    out_dir = ROOT / "reference_style"
    out_dir.mkdir(exist_ok=True)
    (out_dir / "style_profile.json").write_text(json.dumps(profile, indent=2), encoding="utf-8")
    (out_dir / "style_profile.md").write_text(to_markdown(profile), encoding="utf-8")
    print(f"Wrote {out_dir / 'style_profile.json'}")
    print(f"Wrote {out_dir / 'style_profile.md'}")
