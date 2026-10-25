"""
Render a lecture slide-spec (JSON, produced by the Teacher + Academic-PPTX
planning stages) into a fully editable .pptx using python-pptx.

Usage:
    .venv\\Scripts\\python scripts\\build_pptx.py work\\<unit-slug>\\slide_spec.json output\\<unit-slug>.pptx

Slide spec schema: see work/_TEMPLATE/slide_spec.example.json
"""
import json
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent

DEFAULT_THEME = {
    "slide_w_in": 13.333,
    "slide_h_in": 7.5,
    "font_title": "Segoe UI",
    "font_body": "Segoe UI",
    # Classroom-projector floor (house rule, not a per-slide choice): nothing on
    # any slide renders below 24pt. Bands: title 34-40, headings 28-32,
    # body 26-30, labels 24-26, code 28-32. See lecture-pptx-structure/SKILL.md.
    "size_title": 36,
    "size_section": 40,
    "size_heading": 30,
    "size_body": 28,
    "size_small": 26,
    "size_code": 30,
    "size_caption": 24,
    "color_primary": "0B3D5C",
    "color_accent": "E08E1D",
    "color_text": "1A1A1A",
    "color_bg": "FFFFFF",
    "color_emphasis": "C1440E",
    # Token/category palette (Keywords, Identifiers, Constants, Strings,
    # Operators, Special Symbols) -- distinct, colorblind-considerate hues used
    # consistently across token-teaching slides.
    "token_colors": {
        "Keyword": "0B3D5C", "Keywords": "0B3D5C",
        "Identifier": "1F7A5C", "Identifiers": "1F7A5C",
        "Constant": "B8860B", "Constants": "B8860B", "Constant / Literal": "B8860B",
        "String": "6A4C93", "Strings": "6A4C93",
        "Operator": "C1440E", "Operators": "C1440E",
        "Special Symbol": "455A64", "Special Symbols": "455A64", "Special symbol": "455A64",
    },
}


def hexc(h):
    return RGBColor.from_string(h)


class Deck:
    def __init__(self, theme=None):
        self.theme = {**DEFAULT_THEME, **(theme or {})}
        self.prs = Presentation()
        self.prs.slide_width = Inches(self.theme["slide_w_in"])
        self.prs.slide_height = Inches(self.theme["slide_h_in"])
        self.blank_layout = self.prs.slide_layouts[6]

    # ---------- primitives ----------
    def new_slide(self):
        return self.prs.slides.add_slide(self.blank_layout)

    def add_textbox(self, slide, x, y, w, h, text, size, bold=False, color=None,
                     align=PP_ALIGN.LEFT, font=None, anchor=None, italic=False):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.auto_size = None
        try:
            tf.margin_left = tf.margin_right = Pt(2)
            tf.margin_top = tf.margin_bottom = Pt(2)
        except Exception:
            pass
        if anchor:
            tf.vertical_anchor = anchor
        lines = text if isinstance(text, list) else [text]
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            run = p.add_run()
            run.text = line
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.italic = italic
            run.font.name = font or self.theme["font_body"]
            run.font.color.rgb = hexc(color or self.theme["color_text"])
        return box

    def add_bullets(self, slide, x, y, w, h, items, size=None, color=None, font=None, space_after=None):
        size = size or self.theme["size_body"]
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        for i, item in enumerate(items):
            level = 0
            txt = item
            if isinstance(item, dict):
                txt = item.get("text", "")
                level = item.get("level", 0)
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.level = min(level, 4)
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = ("• " if level == 0 else "– ") + txt
            run.font.size = Pt(size - level * 2)
            run.font.name = font or self.theme["font_body"]
            run.font.color.rgb = hexc(color or self.theme["color_text"])
            p.space_after = Pt(space_after if space_after is not None else 8)
        return box

    def add_rect(self, slide, x, y, w, h, fill=None, line=None, text=None,
                 text_size=14, text_color=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, bold=False):
        sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        if fill:
            sp.fill.solid()
            sp.fill.fore_color.rgb = hexc(fill)
        else:
            sp.fill.background()
        if line:
            sp.line.color.rgb = hexc(line)
            sp.line.width = Pt(1.25)
        else:
            sp.line.fill.background()
        sp.shadow.inherit = False
        if text is not None:
            tf = sp.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = text
            run.font.size = Pt(text_size)
            run.font.bold = bold
            run.font.color.rgb = hexc(text_color or self.theme["color_text"])
        return sp

    def estimate_lines(self, text, width_in, size_pt, bold=False, monospace=False):
        """Rough word-wrap simulation so callers can size a box to its actual
        content instead of guessing a fixed height -- at the bigger classroom
        font floor, text that used to fit on one line routinely wraps to two,
        and a stale fixed-height box causes real visual overlap/overflow."""
        if monospace:
            avg_char_w_in = 0.60 * size_pt / 72.0
        else:
            avg_char_w_in = (0.56 if bold else 0.50) * size_pt / 72.0
        usable_w = max(width_in - 0.25, 0.5)
        max_chars = max(1, int(usable_w / avg_char_w_in))
        if monospace:
            # code lines can't reflow at spaces the way prose can -- treat as a
            # hard character count so long single tokens/comments are counted
            # honestly instead of assumed to fit.
            return max(1, -(-len(text) // max_chars))  # ceil division
        lines, cur = 1, 0
        for word in text.split():
            wl = len(word) + 1
            if cur + wl > max_chars and cur > 0:
                lines += 1
                cur = wl
            else:
                cur += wl
        return lines

    def text_block_height(self, text, width_in, size_pt, bold=False, pad_in=0.2):
        lines = self.estimate_lines(text, width_in, size_pt, bold)
        line_h_in = (size_pt * 1.3) / 72.0
        return lines * line_h_in + pad_in

    def add_connector(self, slide, x1, y1, x2, y2, color=None):
        conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        conn.line.color.rgb = hexc(color or self.theme["color_accent"])
        conn.line.width = Pt(1.75)
        line = conn.line._get_or_add_ln()
        tail = line.makeelement(qn('a:tailEnd'), {'type': 'triangle'})
        line.append(tail)
        return conn

    def add_logo_placeholder(self, slide):
        """Top-right logo slot. No real logo supplied yet -- a clearly-marked
        placeholder the user can swap for an actual image later (see
        MSO_SHAPE.PICTURE substitution note in lecture-pptx-build/SKILL.md)."""
        w = self.theme["slide_w_in"]
        logo_w, logo_h = 1.7, 0.8
        x, y = w - 0.5 - logo_w, 0.32
        self.add_rect(slide, x, y, logo_w, logo_h, fill="F2F4F6", line="B8C0C8",
                      text="University Logo", text_size=self.theme["size_caption"],
                      text_color="5A6672", bold=True, shape=MSO_SHAPE.RECTANGLE)
        return x  # left edge, so callers can size the title to stop before it

    def title_bar(self, slide, title, size=None, show_logo=False):
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        title_w = w - 1.0
        if show_logo:
            logo_x = self.add_logo_placeholder(slide)
            title_w = logo_x - 0.5 - 0.3
        title_size = size or self.theme["size_title"]
        box_h = max(0.9, self.text_block_height(title, title_w, title_size, bold=True))
        self.add_textbox(slide, 0.5, 0.3, title_w, box_h, title,
                          title_size, bold=True,
                          color=self.theme["color_primary"], font=self.theme["font_title"])
        return 0.3 + box_h  # bottom edge -- callers should start content at this + a gap

    def footer(self, slide, text):
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        self.add_textbox(slide, 0.5, h - 0.48, w - 1.0, 0.4, text,
                          self.theme["size_caption"], color="6E6E6E", align=PP_ALIGN.LEFT)

    def add_notes(self, slide, text):
        if not text:
            return
        notes = slide.notes_slide
        notes.notes_text_frame.text = text if isinstance(text, str) else "\n".join(text)

    # ---------- layouts ----------
    def layout_title(self, spec):
        slide = self.new_slide()
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        self.add_rect(slide, 0, 0, w, h, fill=self.theme["color_primary"], shape=MSO_SHAPE.RECTANGLE)
        self.add_textbox(slide, 0.75, h / 2 - 1.1, w - 1.5, 1.3, spec["title"],
                          self.theme["size_section"] + 6, bold=True, color="FFFFFF",
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if spec.get("subtitle"):
            self.add_textbox(slide, 0.75, h / 2 + 0.3, w - 1.5, 0.8, spec["subtitle"],
                              self.theme["size_body"], color="D9E2F3", align=PP_ALIGN.CENTER)
        self.add_notes(slide, spec.get("notes"))
        return slide

    def layout_section(self, spec):
        slide = self.new_slide()
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        self.add_rect(slide, 0, 0, w, h, fill=self.theme["color_accent"], shape=MSO_SHAPE.RECTANGLE)
        self.add_textbox(slide, 0.75, h / 2 - 0.6, w - 1.5, 1.2, spec["title"],
                          self.theme["size_section"], bold=True, color="FFFFFF",
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        self.add_notes(slide, spec.get("notes"))
        return slide

    def layout_content(self, spec):
        slide = self.new_slide()
        title_bottom = self.title_bar(slide, spec["title"])
        w = self.theme["slide_w_in"]
        top = title_bottom + 0.2
        self.add_bullets(slide, 0.6, top, w - 1.2, self.theme["slide_h_in"] - top - 0.6, spec.get("bullets", []))
        if spec.get("footer"):
            self.footer(slide, spec["footer"])
        self.add_notes(slide, spec.get("notes"))
        return slide

    def layout_two_column(self, spec):
        slide = self.new_slide()
        title_bottom = self.title_bar(slide, spec["title"])
        w = self.theme["slide_w_in"]
        col_w = (w - 1.6) / 2
        top = title_bottom + 0.15
        left, right = spec.get("left", {}), spec.get("right", {})
        if left.get("heading"):
            self.add_textbox(slide, 0.6, top, col_w, 0.45, left["heading"], self.theme["size_body"],
                              bold=True, color=self.theme["color_accent"])
        if right.get("heading"):
            self.add_textbox(slide, 1.0 + col_w, top, col_w, 0.45, right["heading"], self.theme["size_body"],
                              bold=True, color=self.theme["color_accent"])
        body_top = top + 0.5
        self.add_bullets(slide, 0.6, body_top, col_w, self.theme["slide_h_in"] - body_top - 0.4, left.get("bullets", []),
                          size=self.theme["size_small"])
        self.add_bullets(slide, 1.0 + col_w, body_top, col_w, self.theme["slide_h_in"] - body_top - 0.4, right.get("bullets", []),
                          size=self.theme["size_small"])
        self.add_notes(slide, spec.get("notes"))
        return slide

    def layout_comparison(self, spec):
        slide = self.new_slide()
        title_bottom = self.title_bar(slide, spec["title"])
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        col_w = (w - 1.6) / 2
        top = title_bottom + 0.15
        head_h = 0.6
        self.add_rect(slide, 0.6, top, col_w, head_h, fill=self.theme["color_primary"],
                      text=spec.get("left_head", ""), text_size=self.theme["size_small"], text_color="FFFFFF", bold=True)
        self.add_rect(slide, 1.0 + col_w, top, col_w, head_h, fill=self.theme["color_accent"],
                      text=spec.get("right_head", ""), text_size=self.theme["size_small"], text_color="FFFFFF", bold=True)
        body_top = top + head_h + 0.2
        self.add_bullets(slide, 0.6, body_top, col_w, h - body_top - 0.4, spec.get("left_items", []), size=self.theme["size_small"])
        self.add_bullets(slide, 1.0 + col_w, body_top, col_w, h - body_top - 0.4, spec.get("right_items", []), size=self.theme["size_small"])
        self.add_notes(slide, spec.get("notes"))
        return slide

    def layout_diagram(self, spec):
        """Native, fully-editable flow diagram: sequence of boxes with arrows,
        left-to-right or top-to-bottom depending on node count."""
        slide = self.new_slide()
        title_bottom = self.title_bar(slide, spec["title"])
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        nodes = spec["diagram"]["nodes"]
        sublabels = spec["diagram"].get("sublabels")
        direction = spec["diagram"].get("direction", "LR")
        intro = spec["diagram"].get("intro")
        n = len(nodes)
        top = title_bottom + 0.2
        if intro:
            intro_h = self.text_block_height(intro, w - 1.2, self.theme["size_small"])
            self.add_textbox(slide, 0.6, top, w - 1.2, intro_h, intro, self.theme["size_small"],
                              color=self.theme["color_text"], italic=True)
            top += intro_h + 0.2

        caption = spec.get("caption")
        cap_reserve = (self.text_block_height(caption, w - 1.0, self.theme["size_caption"]) + 0.25) if caption else 0.3
        avail_h = h - top - cap_reserve

        if direction == "LR":
            margin, gap = 0.4, 0.25
            box_w = (w - 2 * margin - gap * (n - 1)) / n
            if sublabels:
                max_lab_lines = max(self.estimate_lines(lbl, box_w, self.theme["size_caption"]) for lbl in sublabels)
                sub_reserve = max_lab_lines * (self.theme["size_caption"] * 1.3 / 72.0) + 0.2 + 0.15
            else:
                sub_reserve = 0
            box_h = min(2.3, avail_h - sub_reserve)
            y = top + (avail_h - box_h - sub_reserve) / 2
            xs = []
            x = margin
            for node in nodes:
                self.add_rect(slide, x, y, box_w, box_h, fill=self.theme["color_accent"],
                              text=node, text_size=self.theme["size_small"], text_color="FFFFFF", bold=True)
                xs.append(x)
                x += box_w + gap
            for i in range(n - 1):
                x1 = xs[i] + box_w
                x2 = xs[i + 1]
                self.add_connector(slide, x1, y + box_h / 2, x2, y + box_h / 2)
            if sublabels:
                x = margin
                lab_h = max_lab_lines * (self.theme["size_caption"] * 1.3 / 72.0) + 0.2
                for label in sublabels:
                    self.add_textbox(slide, x, y + box_h + 0.15, box_w, lab_h, label,
                                      self.theme["size_caption"], color="6E6E6E",
                                      align=PP_ALIGN.CENTER)
                    x += box_w + gap
        else:
            gap = 0.25
            box_h = min(0.85, (avail_h - gap * (n - 1)) / n)
            box_w = w - 2.4
            x = 1.2
            y = top
            ys = []
            for node in nodes:
                self.add_rect(slide, x, y, box_w, box_h, fill=self.theme["color_accent"],
                              text=node, text_size=self.theme["size_small"], text_color="FFFFFF", bold=True)
                ys.append(y)
                y += box_h + gap
            for i in range(n - 1):
                y1 = ys[i] + box_h
                y2 = ys[i + 1]
                self.add_connector(slide, x + box_w / 2, y1, x + box_w / 2, y2)
        if caption:
            cap_h = self.text_block_height(caption, w - 1.0, self.theme["size_caption"])
            self.add_textbox(slide, 0.5, h - 0.15 - cap_h, w - 1.0, cap_h, caption,
                              self.theme["size_caption"], color="6E6E6E")
        self.add_notes(slide, spec.get("notes"))
        return slide

    def layout_table(self, spec):
        slide = self.new_slide()
        title_bottom = self.title_bar(slide, spec["title"])
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        headers = spec["headers"]
        rows = spec["rows"]
        nrows, ncols = len(rows) + 1, len(headers)
        top = title_bottom + 0.2
        insight_h = self.text_block_height(spec["insight"], w - 1.2, self.theme["size_body"]) if spec.get("insight") else 0
        example_h = (max(1.0, self.estimate_lines(spec["example"], w - 1.4, self.theme["size_code"], monospace=True) * 0.55 + 0.3)
                      if spec.get("example") else 0)
        # PowerPoint auto-expands table rows to fit wrapped cell text regardless
        # of the height we request on add_table -- so compute the height it will
        # actually render at (per-row, from real content) instead of picking a
        # number and hoping the table compresses to match. If the honest total
        # doesn't fit, that's a genuine overflow for the caller to fix by
        # shortening content or splitting the slide, not something to paper over.
        col_w = (w - 1.2) / ncols
        header_h = max(0.55, max(self.text_block_height(str(ht), col_w, self.theme["size_body"], bold=True)
                                  for ht in headers))
        row_heights = [max(0.5, max(self.text_block_height(str(v), col_w, self.theme["size_small"]) for v in row))
                       for row in rows]
        tbl_h = header_h + sum(row_heights)
        gfx = slide.shapes.add_table(nrows, ncols, Inches(0.6), Inches(top), Inches(w - 1.2), Inches(tbl_h))
        gfx.table.rows[0].height = Emu(int(Inches(header_h)))
        for r, rh in enumerate(row_heights, start=1):
            gfx.table.rows[r].height = Emu(int(Inches(rh)))
        table = gfx.table
        for c, htext in enumerate(headers):
            cell = table.cell(0, c)
            cell.text = str(htext)
            cell.fill.solid()
            cell.fill.fore_color.rgb = hexc(self.theme["color_primary"])
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for r in p.runs:
                    r.font.bold = True
                    r.font.color.rgb = hexc("FFFFFF")
                    r.font.size = Pt(self.theme["size_body"])
        for r, row in enumerate(rows, start=1):
            for c, val in enumerate(row):
                cell = table.cell(r, c)
                cell.text = str(val)
                for p in cell.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(self.theme["size_small"])
        y = top + tbl_h + 0.2
        if spec.get("insight"):
            self.add_textbox(slide, 0.6, y, w - 1.2, insight_h, spec["insight"],
                              self.theme["size_body"], italic=True, color=self.theme["color_emphasis"])
            y += insight_h + 0.2
        if spec.get("example"):
            ex = spec["example"]
            box = slide.shapes.add_textbox(Inches(0.6), Inches(y), Inches(w - 1.2), Inches(example_h))
            box.fill.solid()
            box.fill.fore_color.rgb = hexc("1E1E1E")
            tf = box.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = ex
            run.font.name = "Consolas"
            run.font.size = Pt(self.theme["size_code"])
            run.font.color.rgb = hexc("D4D4D4")
        self.add_notes(slide, spec.get("notes"))
        return slide

    def layout_image(self, spec):
        slide = self.new_slide()
        title_bottom = self.title_bar(slide, spec["title"])
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        img_path = spec["image_path"]
        top = title_bottom + 0.2
        cap_h = self.text_block_height(spec["caption"], w - 1.0, self.theme["size_caption"]) + 0.2 if spec.get("caption") else 0.2
        max_h = h - top - cap_h
        pic = slide.shapes.add_picture(img_path, Inches(0.6), Inches(top), height=Inches(max_h))
        if pic.width > Inches(w - 1.2):
            ratio = Inches(w - 1.2) / pic.width
            pic.width = Inches(w - 1.2)
            pic.height = int(pic.height * ratio)
        pic.left = int((self.prs.slide_width - pic.width) / 2)
        if spec.get("caption"):
            self.footer(slide, spec["caption"])
        self.add_notes(slide, spec.get("notes"))
        return slide

    def layout_formula(self, spec):
        """Formula rendered via matplotlib mathtext -> PNG (no LaTeX install needed),
        embedded as a picture. Plain-text formula also kept in speaker notes for editability."""
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        slide = self.new_slide()
        title_bottom = self.title_bar(slide, spec["title"])
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        top = title_bottom + 0.3

        tmp_dir = ROOT / ".cache" / "tmp"
        tmp_dir.mkdir(parents=True, exist_ok=True)
        img_path = tmp_dir / f"formula_{abs(hash(spec['formula']))}.png"
        fig = plt.figure(figsize=(6, 1.4), dpi=300)
        fig.patch.set_alpha(0)
        fig.text(0.5, 0.5, f"${spec['formula']}$", fontsize=26, ha="center", va="center",
                  color=f"#{self.theme['color_primary']}")
        fig.savefig(img_path, transparent=True, bbox_inches="tight", pad_inches=0.1)
        plt.close(fig)

        pic = slide.shapes.add_picture(str(img_path), Inches(1.0), Inches(top), width=Inches(w - 2.0))
        pic.left = int((self.prs.slide_width - pic.width) / 2)

        expl = spec.get("explanation", [])
        if expl:
            expl_top = top + pic.height / 914400 + 0.3
            self.add_bullets(slide, 0.8, expl_top, w - 1.6,
                              h - expl_top - 0.3, expl, size=self.theme["size_small"])
        notes = list(spec.get("notes", []) or [])
        notes.append(f"Formula (plain text, for editing): {spec['formula']}")
        self.add_notes(slide, notes)
        return slide

    def layout_code(self, spec):
        slide = self.new_slide()
        title_bottom = self.title_bar(slide, spec["title"])
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        lines = spec["code"].split("\n")
        line_h = (self.theme["size_code"] + 10) / 72.0
        # count wrapped physical lines per source line, not just len(lines) --
        # a long inline comment at the 28-32pt code floor routinely wraps, and
        # under-counting here caused the code box to overflow both above (into
        # the title) and below (behind the next element) in practice.
        wrapped_line_count = sum(
            self.estimate_lines(line if line.strip() else " ", w - 1.4, self.theme["size_code"], monospace=True)
            for line in lines
        )
        code_h = max(1.0, wrapped_line_count * line_h + 0.3)
        top = title_bottom + 0.2
        box = slide.shapes.add_textbox(Inches(0.6), Inches(top), Inches(w - 1.2), Inches(code_h))
        box.fill.solid()
        box.fill.fore_color.rgb = hexc("1E1E1E")
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.25)
        tf.margin_top = Inches(0.15)
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            run = p.add_run()
            run.text = line if line.strip() else " "
            run.font.name = "Consolas"
            run.font.size = Pt(self.theme["size_code"])
            run.font.color.rgb = hexc("D4D4D4")
        y = top + code_h + 0.3
        if spec.get("output"):
            self.add_textbox(slide, 0.6, y, 2.4, 0.45, "Output:", self.theme["size_small"], bold=True,
                              color=self.theme["color_accent"])
            out_box = slide.shapes.add_textbox(Inches(0.6), Inches(y + 0.45), Inches(w - 1.2), Inches(1.0))
            out_box.fill.solid()
            out_box.fill.fore_color.rgb = hexc("F2F2F2")
            out_box.line.color.rgb = hexc("CCCCCC")
            out_box.line.width = Pt(1)
            otf = out_box.text_frame
            otf.word_wrap = True
            otf.margin_left = Inches(0.2)
            for i, line in enumerate(spec["output"].split("\n")):
                p = otf.paragraphs[0] if i == 0 else otf.add_paragraph()
                run = p.add_run()
                run.text = line
                run.font.name = "Consolas"
                run.font.size = Pt(self.theme["size_code"])
                run.font.color.rgb = hexc("202020")
            y += 0.45 + 1.0 + 0.25
        if spec.get("annotation"):
            # No bottom-clamp here on purpose: silently clamping the y-position
            # when content runs long causes it to overlap the previous element
            # instead of visibly overflowing -- let qa_check.py's bounds check
            # catch a genuine overflow so the spec gets simplified/split instead.
            ann_h = self.text_block_height(spec["annotation"], w - 1.2, self.theme["size_small"])
            self.add_textbox(slide, 0.6, y, w - 1.2, ann_h, spec["annotation"],
                              self.theme["size_small"], italic=True, color="6E6E6E")
        self.add_notes(slide, spec.get("notes"))
        return slide

    def layout_recap(self, spec):
        slide = self.new_slide()
        title_bottom = self.title_bar(slide, spec.get("title", "Recap Questions"))
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        items = [f"{i+1}.  {q}" for i, q in enumerate(spec["questions"])]
        closing = spec.get("closing")
        closing_h = (self.text_block_height(closing, w - 1.6, self.theme["size_small"], bold=True) + 0.3
                     if closing else 0)
        top = title_bottom + 0.2
        body_bottom = h - 0.3 - (closing_h + 0.2 if closing else 0)
        self.add_bullets(slide, 0.6, top, w - 1.2, body_bottom - top, items,
                          size=self.theme["size_body"], space_after=10)
        if closing:
            self.add_rect(slide, 0.6, h - 0.3 - closing_h, w - 1.2, closing_h, fill=self.theme["color_primary"],
                          text=closing, text_size=self.theme["size_small"], text_color="FFFFFF", bold=True,
                          shape=MSO_SHAPE.RECTANGLE)
        self.add_notes(slide, spec.get("notes"))
        return slide

    def layout_mistake(self, spec):
        """Common-mistake callout slide."""
        slide = self.new_slide()
        title_bottom = self.title_bar(slide, spec.get("title", "Common Mistake"))
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        top = title_bottom + 0.2
        mistake_h = max(1.2, self.text_block_height(spec.get("mistake", ""), w - 1.6, self.theme["size_body"], bold=True) + 0.4)
        self.add_rect(slide, 0.6, top, w - 1.2, mistake_h, fill="FCE4E4", line="C00000",
                      text=spec.get("mistake", ""), text_size=self.theme["size_body"], text_color="C00000", bold=True)
        if spec.get("fix"):
            label_top = top + mistake_h + 0.25
            self.add_textbox(slide, 0.6, label_top, w - 1.2, 0.5, "Correct approach:", self.theme["size_small"], bold=True,
                              color=self.theme["color_primary"])
            self.add_bullets(slide, 0.6, label_top + 0.55, w - 1.2, h - (label_top + 0.55) - 0.3, spec["fix"], size=self.theme["size_small"])
        self.add_notes(slide, spec.get("notes"))
        return slide

    def layout_token_categories(self, spec):
        """Slide 1/2 of a token-teaching pair: definition + the category set,
        as clean bordered pills (no filled color blocks) so it reads as
        classification, not decoration."""
        slide = self.new_slide()
        title_bottom = self.title_bar(slide, spec["title"], show_logo=spec.get("show_logo", True))
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        colors = self.theme["token_colors"]

        top = title_bottom + 0.2
        if spec.get("definition_label") or spec.get("definition"):
            box_h = 1.5
            box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(top),
                                          Inches(w - 1.2), Inches(box_h))
            box.fill.solid()
            box.fill.fore_color.rgb = hexc("F2F4F6")
            box.line.color.rgb = hexc("D5DBE0")
            box.line.width = Pt(1)
            box.shadow.inherit = False
            tf = box.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = Inches(0.3)
            tf.margin_top = Inches(0.18)
            p0 = tf.paragraphs[0]
            r0 = p0.add_run()
            r0.text = spec.get("definition_label", "DEFINITION")
            r0.font.size = Pt(self.theme["size_small"])
            r0.font.bold = True
            r0.font.color.rgb = hexc(self.theme["color_accent"])
            p1 = tf.add_paragraph()
            p1.space_before = Pt(6)
            r1 = p1.add_run()
            r1.text = spec.get("definition", "")
            r1.font.size = Pt(self.theme["size_body"])
            r1.font.color.rgb = hexc(self.theme["color_text"])
            top += box_h + 0.3

        # 3 columns x 2 rows -- a single row of 6 is too narrow per pill for
        # words like "Identifiers"/"Operators" at classroom-readable font size
        # and forces an ugly mid-word wrap.
        categories = spec.get("categories", list(colors.keys()))
        n = len(categories)
        cols = 3
        rows = (n + cols - 1) // cols
        gap_x, gap_y = 0.35, 0.25
        pill_w = (w - 1.2 - gap_x * (cols - 1)) / cols
        pill_h = 1.0
        for i, cat in enumerate(categories):
            r, c_idx = divmod(i, cols)
            x = 0.6 + c_idx * (pill_w + gap_x)
            y = top + r * (pill_h + gap_y)
            c = colors.get(cat, self.theme["color_primary"])
            pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y),
                                           Inches(pill_w), Inches(pill_h))
            pill.adjustments[0] = 0.12
            pill.fill.solid()
            pill.fill.fore_color.rgb = hexc("FFFFFF")
            pill.line.color.rgb = hexc(c)
            pill.line.width = Pt(2)
            pill.shadow.inherit = False
            tf = pill.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = cat
            run.font.size = Pt(self.theme["size_small"])
            run.font.bold = True
            run.font.color.rgb = hexc(c)
        top += rows * pill_h + (rows - 1) * gap_y + 0.4

        if spec.get("lead_in"):
            self.add_textbox(slide, 0.6, top, w - 1.2, 0.6, spec["lead_in"],
                              self.theme["size_body"], color=self.theme["color_text"],
                              align=PP_ALIGN.CENTER)
        if spec.get("footer"):
            self.footer(slide, spec["footer"])
        self.add_notes(slide, spec.get("notes"))
        return slide

    def layout_token_breakdown(self, spec):
        """Slide 2/2: one worked statement, color-coded per token, broken into
        labeled cards. No dark filled token blocks -- light cards with a thin
        colored accent bar carry the category color instead."""
        slide = self.new_slide()
        title_bottom = self.title_bar(slide, spec["title"], show_logo=spec.get("show_logo", True))
        w, h = self.theme["slide_w_in"], self.theme["slide_h_in"]
        colors = self.theme["token_colors"]
        top = title_bottom + 0.2

        statement = spec["statement"]  # list of {"text":.., "category":..}
        stmt_h = 1.15
        stmt_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(top),
                                           Inches(w - 1.2), Inches(stmt_h))
        stmt_box.fill.solid()
        stmt_box.fill.fore_color.rgb = hexc("FFFFFF")
        stmt_box.line.color.rgb = hexc("D5DBE0")
        stmt_box.line.width = Pt(1.5)
        stmt_box.shadow.inherit = False
        tf = stmt_box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        for i, tok in enumerate(statement):
            run = p.add_run()
            run.text = tok["text"] + (" " if i < len(statement) - 1 else "")
            run.font.name = "Consolas"
            run.font.size = Pt(self.theme["size_code"] + 2)
            run.font.bold = True
            run.font.color.rgb = hexc(colors.get(tok["category"], self.theme["color_text"]))
        top += stmt_h + 0.4

        cards = spec.get("cards", [])
        n = len(cards)
        gap = 0.25
        card_w = (w - 1.2 - gap * (n - 1)) / n
        card_h = 1.75
        accent_h = 0.1
        x = 0.6
        for card in cards:
            c = colors.get(card["category"], self.theme["color_primary"])
            self.add_rect(slide, x, top, card_w, accent_h, fill=c, line=None, shape=MSO_SHAPE.RECTANGLE)
            body = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(top + accent_h),
                                           Inches(card_w), Inches(card_h - accent_h))
            body.fill.solid()
            body.fill.fore_color.rgb = hexc("FAFBFC")
            body.line.color.rgb = hexc("D5DBE0")
            body.line.width = Pt(1)
            body.shadow.inherit = False
            tf = body.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p1 = tf.paragraphs[0]
            p1.alignment = PP_ALIGN.CENTER
            r1 = p1.add_run()
            r1.text = card["token"]
            r1.font.name = "Consolas"
            r1.font.size = Pt(self.theme["size_code"])
            r1.font.bold = True
            r1.font.color.rgb = hexc(c)
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(10)
            r2 = p2.add_run()
            r2.text = card["category"]
            r2.font.size = Pt(self.theme["size_small"])
            r2.font.color.rgb = hexc(self.theme["color_text"])
            x += card_w + gap
        top += card_h + 0.35

        if spec.get("insight"):
            self.add_textbox(slide, 0.6, top, w - 1.2, 0.9, spec["insight"],
                              self.theme["size_body"], italic=True, color=self.theme["color_emphasis"],
                              align=PP_ALIGN.CENTER)
        if spec.get("footer"):
            self.footer(slide, spec["footer"])
        self.add_notes(slide, spec.get("notes"))
        return slide

    LAYOUT_MAP = {}

    def build(self, spec: dict):
        self.theme = {**self.theme, **spec.get("meta", {}).get("theme", {})}
        self.prs.slide_width = Inches(self.theme["slide_w_in"])
        self.prs.slide_height = Inches(self.theme["slide_h_in"])
        mapping = {
            "title": self.layout_title,
            "section": self.layout_section,
            "content": self.layout_content,
            "two_column": self.layout_two_column,
            "comparison": self.layout_comparison,
            "diagram": self.layout_diagram,
            "table": self.layout_table,
            "image": self.layout_image,
            "formula": self.layout_formula,
            "code": self.layout_code,
            "recap": self.layout_recap,
            "mistake": self.layout_mistake,
            "token_categories": self.layout_token_categories,
            "token_breakdown": self.layout_token_breakdown,
        }
        for i, slide_spec in enumerate(spec["slides"]):
            layout = slide_spec.get("layout")
            fn = mapping.get(layout)
            if not fn:
                raise ValueError(f"Slide {i}: unknown layout '{layout}'")
            fn(slide_spec)
        return self.prs


def main():
    if len(sys.argv) < 3:
        print("Usage: build_pptx.py <slide_spec.json> <output.pptx>")
        sys.exit(1)
    spec_path = Path(sys.argv[1])
    out_path = Path(sys.argv[2])
    out_path.parent.mkdir(parents=True, exist_ok=True)
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    deck = Deck()
    prs = deck.build(spec)
    prs.save(str(out_path))
    print(f"Wrote {out_path} ({len(spec['slides'])} slides)")


if __name__ == "__main__":
    main()
