import copy
import io
import json
import math
import os
import re
import shutil
import subprocess
import tempfile
import threading
import uuid
from datetime import datetime
from pathlib import Path

from . import blob_store

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "templates"
SHARED_DIR = TEMPLATES_DIR / "_shared"
DATA_DIR = Path(__file__).resolve().parent.parent / "data"
SIGNATURES_DIR = DATA_DIR / "signatures"
TEMP_SIGNATURES_DIR = DATA_DIR / "temp_signatures"
if not os.environ.get("VERCEL"):
    # These are only ever written to on a normal persistent filesystem -- Vercel's
    # deployment bundle is read-only (signature storage goes through Blob instead,
    # see signature_path/temp_signature_path below, once BLOB_READ_WRITE_TOKEN is set),
    # so these mkdir calls would fail there even before Blob is configured.
    SIGNATURES_DIR.mkdir(parents=True, exist_ok=True)
    TEMP_SIGNATURES_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR = Path(tempfile.gettempdir()) / "quickprop_output" if os.environ.get("VERCEL") else Path(__file__).resolve().parent.parent / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

IMAGE_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def list_styles() -> list[str]:
    return sorted(p.name for p in SHARED_DIR.iterdir() if p.is_dir())


def load_schema(style: str) -> dict:
    path = SHARED_DIR / style / "schema.json"
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_companies() -> dict:
    with open(DATA_DIR / "companies.json", "r", encoding="utf-8") as f:
        return json.load(f)


def style_preview_path(style: str) -> Path | None:
    for ext in ("jpg", "jpeg", "png"):
        candidate = SHARED_DIR / style / f"preview.{ext}"
        if candidate.exists():
            return candidate
    return None


def signature_path(company: str) -> bytes | Path | None:
    """The company's saved *default* rep signature (uploaded once via the admin page) --
    used only as a fallback when a specific generation doesn't supply its own rep
    signature (e.g. a different staff member sending this particular proposal).
    Returns raw bytes when Blob storage is active (Vercel), a local Path otherwise --
    _place_signature() below accepts either."""
    if blob_store.BLOB_ENABLED:
        for ext in (".png", ".jpg", ".jpeg"):
            data = blob_store.blob_get(f"signatures/{company}{ext}")
            if data:
                return data
        return None
    for ext in (".png", ".jpg", ".jpeg"):
        candidate = SIGNATURES_DIR / f"{company}{ext}"
        if candidate.exists():
            return candidate
    return None


def temp_signature_path(signature_id: str) -> bytes | Path | None:
    """Resolves a one-off signature uploaded for a single generation (rep or client),
    identified by the id returned from the upload endpoint."""
    if not re.fullmatch(r"[0-9a-fA-F-]{8,64}", signature_id or ""):
        return None
    if blob_store.BLOB_ENABLED:
        for ext in (".png", ".jpg", ".jpeg"):
            data = blob_store.blob_get(f"temp_signatures/{signature_id}{ext}")
            if data:
                return data
        return None
    for ext in (".png", ".jpg", ".jpeg"):
        candidate = TEMP_SIGNATURES_DIR / f"{signature_id}{ext}"
        if candidate.exists():
            return candidate
    return None


def _find_shape(slide, shape_name: str):
    return next(s for s in slide.shapes if s.name == shape_name)


def _try_remove_shape(slide, shape_name: str) -> None:
    try:
        shape = _find_shape(slide, shape_name)
    except StopIteration:
        return
    shape._element.getparent().remove(shape._element)


def _enable_shrink_to_fit(slide, shape, slide_width_emu=None, slide_height_emu=None, badge_colors=None) -> None:
    """PowerPoint's own "shrink text on overflow" only ever recalculates when a human
    types into the box in the app — opening or exporting a file we generated headlessly
    never triggers it, so setting auto_size alone silently does nothing.

    Only touches shapes with exactly one paragraph and one run — i.e. the shape's whole
    job is to hold this one value. A shape that mixes this value into a longer sentence
    (several runs sharing one paragraph) is left alone: forcing word_wrap off there would
    cram the entire surrounding sentence onto one line instead of just the substituted
    span, which would look far worse than the overflow it was meant to fix.

    Many of these shapes are plain, invisible textboxes sitting directly on top of a
    decorative "pill"/badge that's actually baked into the slide's fixed background image
    at a one-line size. An earlier version of this drew a brand-new replacement shape over
    long values instead of just shrinking the text -- but that new shape's size never quite
    matched the pill baked into the background, so the original pill peeked out from behind
    it, looking like two mismatched, stacked pills. Keeping the exact same shape and just
    shrinking the font (wrapping onto a second line within its own footprint only if it
    still doesn't fit at a legible size) guarantees there is only ever the one, original,
    correctly-sized pill on screen.
    """
    if not shape.has_text_frame or not shape.width:
        return
    tf = shape.text_frame
    if len(tf.paragraphs) != 1 or len(tf.paragraphs[0].runs) != 1:
        return
    box_width_in = shape.width / 914400
    box_height_in = shape.height / 914400
    run = tf.paragraphs[0].runs[0]
    text = run.text
    if not text:
        return
    original_pt = (run.font.size.pt if run.font.size else 18)

    def fits_unwrapped(size_pt):
        return len(text) * (size_pt * 0.52) / 72 <= box_width_in * 0.92

    if fits_unwrapped(original_pt):
        tf.word_wrap = False
        return

    def lines_needed(size_pt, width_in):
        char_width_in = size_pt * 0.52 / 72
        chars_per_line = max(1, int(width_in * 0.9 / char_width_in))
        return max(1, math.ceil(len(text) / chars_per_line))

    def block_height_in(size_pt, lines):
        return lines * size_pt * 1.35 / 72

    # Shrink the font (and let it wrap within the shape's own original width) until the
    # wrapped block's height fits back inside the shape's own original height -- never
    # resizing or replacing the shape itself. Floors at 7pt so it stays legible even if a
    # genuinely huge value can't be made to fit cleanly; a little overflow at that point is
    # less visually broken than a second, mismatched pill.
    size_pt = original_pt
    while size_pt > 7:
        lines = lines_needed(size_pt, box_width_in)
        if block_height_in(size_pt, lines) <= box_height_in * 0.92:
            break
        size_pt -= 1

    tf.word_wrap = True
    run.font.size = Pt(size_pt)


def _set_run_text(slide, target: dict, value: str) -> None:
    shape = _find_shape(slide, target["shape"])
    para = shape.text_frame.paragraphs[target["paragraph"]]
    prefix = target.get("prefix", "")
    suffix = target.get("suffix", "")
    if target["run"] == "all":
        # Paragraph is split across many runs (e.g. lorem-ipsum spell-check artifacts).
        # Collapse it to a single run carrying the new text.
        first_run = para.runs[0]
        first_run.text = f"{prefix}{value}{suffix}"
        for extra_run in para.runs[1:]:
            extra_run._r.getparent().remove(extra_run._r)
    else:
        run = para.runs[target["run"]]
        run.text = f"{prefix}{value}{suffix}"
    return shape


def _global_find_replace(prs, replacements: list[tuple[str, str]], badge_colors=None) -> None:
    """Finds literal text anywhere in the deck and replaces it, even when PowerPoint
    (or a "track changes" pass) has split that text across many tiny runs — a common
    artifact in templates bought from marketplaces. Works at the paragraph level: if a
    paragraph's full text contains `find`, the whole paragraph collapses to one run with
    the substitution applied (losing intra-paragraph formatting, keeping the first run's
    font). `replacements` is sorted longest-find-first internally so a shorter find can't
    clobber part of a longer one first (e.g. "Progressive UX" vs "Progressive UX Inc.").
    """
    ordered = sorted(replacements, key=lambda pair: len(pair[0]), reverse=True)
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            changed = False
            for para in shape.text_frame.paragraphs:
                if not para.runs:
                    continue
                full = "".join(r.text for r in para.runs)
                new_full = full
                for find, value in ordered:
                    if find in new_full:
                        new_full = new_full.replace(find, value)
                if new_full != full:
                    para.runs[0].text = new_full
                    for extra in para.runs[1:]:
                        extra._r.getparent().remove(extra._r)
                    changed = True
            # Fixed, single-line signature-block boxes (e.g. "To: {client_name}") have no
            # wrap/shrink protection of their own — a long substituted value would just
            # overflow into neighboring shapes. Route changed shapes through the same
            # shrink-to-fit/badge logic the `fields`/`branding_targets` paths already get.
            # Only for short identity-style values (names/dates) though — a shape holding
            # a full substituted paragraph (e.g. scope of work) is meant to wrap normally
            # within its box, not get wrapped in a little pill-shaped badge.
            first_runs = shape.text_frame.paragraphs[0].runs
            if changed and first_runs and len(first_runs[0].text) <= 80:
                _enable_shrink_to_fit(slide, shape, prs.slide_width, prs.slide_height, badge_colors)


def _remove_paragraphs(slide, shape_name: str, paragraph_indices: list[int]) -> None:
    shape = _find_shape(slide, shape_name)
    paragraphs = shape.text_frame.paragraphs
    for idx in sorted(paragraph_indices, reverse=True):
        p_elem = paragraphs[idx]._p
        p_elem.getparent().remove(p_elem)


def _remove_slide(prs, slide_index_zero_based: int) -> None:
    sld_id_lst = prs.slides._sldIdLst
    entries = list(sld_id_lst)
    entry = entries[slide_index_zero_based]
    # Dropping the relationship (not just the sldId reference) removes the underlying
    # slide part from the package too — otherwise its orphaned partname can collide
    # with one PowerPoint's internal counter later hands out to a freshly-added slide.
    prs.part.drop_rel(entry.get(qn("r:id")))
    sld_id_lst.remove(entry)


def duplicate_slide(prs, source_index: int, insert_after_index: int):
    """Clones the slide at `source_index` (0-based) and inserts the copy immediately
    after `insert_after_index` (0-based). This is how the deck grows to fit however much
    content a given proposal actually has, instead of every template being limited to a
    fixed number of slots. Returns the new slide."""
    source = prs.slides[source_index]
    dest = prs.slides.add_slide(source.slide_layout)

    for shp in list(dest.shapes):
        shp._element.getparent().remove(shp._element)

    for shape in source.shapes:
        new_el = copy.deepcopy(shape._element)
        for blip in new_el.iter(qn("a:blip")):
            rId = blip.get(qn("r:embed"))
            if rId and rId in source.part.rels:
                image_part = source.part.rels[rId].target_part
                new_rId = dest.part.relate_to(image_part, IMAGE_RELTYPE)
                blip.set(qn("r:embed"), new_rId)
        dest.shapes._spTree.append(new_el)

    sld_id_lst = prs.slides._sldIdLst
    entries = list(sld_id_lst)
    new_entry = entries[-1]
    sld_id_lst.remove(new_entry)
    sld_id_lst.insert(insert_after_index + 1, new_entry)
    return dest


def insert_blank_slide(prs, insert_after_index: int, layout_index: int = -1):
    """Adds a genuinely empty slide (from a blank layout, not a clone of existing content)
    and inserts it after `insert_after_index` (0-based). Useful when a template has no
    spare slide and every existing slide is too busy (title art, dense legal text) to
    safely clone as a stencil for something like an inserted pricing table."""
    layout = prs.slide_layouts[layout_index]
    dest = prs.slides.add_slide(layout)
    for shp in list(dest.shapes):
        shp._element.getparent().remove(shp._element)

    sld_id_lst = prs.slides._sldIdLst
    entries = list(sld_id_lst)
    new_entry = entries[-1]
    sld_id_lst.remove(new_entry)
    sld_id_lst.insert(insert_after_index + 1, new_entry)
    return dest


def _strip_table_style(table) -> None:
    """Removes the default blue/banded PowerPoint table theme so cells render as
    plain white boxes with dark text — matching the rest of the proposal instead of
    standing out as an unstyled auto-generated block."""
    table.first_row = False
    table.horz_banding = False
    for row in table.rows:
        for cell in row.cells:
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)


def _capture_text_style(slide, shape_name: str, paragraph_index: int = 0) -> dict:
    """Reads the font used by an existing heading/body textbox in the template, so
    text we generate programmatically for the content flow matches it exactly."""
    shape = _find_shape(slide, shape_name)
    run = shape.text_frame.paragraphs[paragraph_index].runs[0]
    color = None
    try:
        if run.font.color and run.font.color.type is not None:
            color = run.font.color.rgb
    except AttributeError:
        color = None
    return {
        "font_name": run.font.name,
        "size": run.font.size,
        "bold": run.font.bold,
        "color": color,
    }


def _apply_text_style(run, style: dict) -> None:
    if style.get("font_name"):
        run.font.name = style["font_name"]
    if style.get("size"):
        run.font.size = style["size"]
    if style.get("bold") is not None:
        run.font.bold = style["bold"]
    run.font.color.rgb = style.get("color") or RGBColor(0x1A, 0x1A, 0x1A)


def _estimate_text_height(text: str, box_width_in: float, font_size_pt: float, min_lines: int = 1) -> float:
    """Estimates how tall a wrapped block of text will render, driven by the box's
    actual width and the actual font size — a fixed "chars per line" guess badly
    underestimates height once the font is bigger than the guess assumed, which is
    exactly what caused blocks to overlap the next heading further down the page."""
    char_width_in = font_size_pt * 0.52 / 72
    chars_per_line = max(1, int(box_width_in / char_width_in))
    lines = max(min_lines, math.ceil(len(text) / chars_per_line)) if text else min_lines
    line_height_in = font_size_pt * 1.35 / 72
    return lines * line_height_in * 1.15  # 15% safety margin for estimation error


class ContentFlow:
    """Packs a sequence of heading/body blocks and tables onto slides, stacking each
    new block directly under the last one and only starting a fresh (cloned) slide once
    the current one actually runs out of room — so a proposal with 2 sections and one
    with 6 both render as "just enough" pages, no fixed one-block-per-slide layout."""

    def __init__(self, prs, layout: dict, heading_style: dict, body_style: dict):
        self.prs = prs
        self.layout = layout
        self.heading_style = heading_style
        self.body_style = body_style
        self.left = Inches(layout["left_in"])
        self.width = Inches(layout["width_in"])
        self.top_start = Inches(layout["top_start_in"])
        self.bottom_limit = Inches(layout["bottom_limit_in"])
        self.gap = Inches(layout.get("gap_in", 0.25))

        self.first_slide_index = layout["first_slide"] - 1
        first_slide = prs.slides[self.first_slide_index]

        # Any later page is cloned from a PRISTINE stencil, never from the slide that's
        # actively accumulating this run's content — cloning an in-progress page would
        # carry every block already placed on it into the "new" page too.
        last_index = len(prs.slides) - 1
        duplicate_slide(prs, self.first_slide_index, last_index)
        self.pristine_index = last_index + 1
        pristine = prs.slides[self.pristine_index]
        for shape_name in layout["strip_shapes"]:
            _try_remove_shape(pristine, shape_name)

        for shape_name in layout["strip_shapes"]:
            _try_remove_shape(first_slide, shape_name)

        self.current_slide = first_slide
        self.current_top = self.top_start
        self.insert_at = self.first_slide_index
        self.pages_used = 1

    def _next_page(self):
        self.current_slide = duplicate_slide(self.prs, self.pristine_index, self.insert_at)
        self.insert_at += 1
        self.pristine_index += 1  # the stencil itself just shifted right by one
        self.current_top = self.top_start
        self.pages_used += 1

    def finish(self):
        """Removes the pristine stencil slide once packing is done — it was only ever
        scratch material for cloning, never meant to appear in the final deck."""
        _remove_slide(self.prs, self.pristine_index)

    def _place(self, height_emu: int, add_fn):
        if self.current_top + height_emu > self.bottom_limit and self.current_top > self.top_start:
            self._next_page()
        add_fn(self.current_slide, self.current_top)
        self.current_top += height_emu + self.gap

    def add_text_block(self, heading: str, body: str):
        heading_font_pt = (self.heading_style.get("size").pt if self.heading_style.get("size") else 18)
        heading_h = Inches(heading_font_pt * 1.4 / 72 + 0.08) if heading else 0
        body_font_pt = (self.body_style.get("size").pt if self.body_style.get("size") else 14)
        width_in = self.layout["width_in"]
        body_h = Inches(_estimate_text_height(body, width_in, body_font_pt))
        gap_after_heading = Inches(0.05) if heading else 0
        total_h = heading_h + body_h

        def add(slide, top):
            if heading:
                h_box = slide.shapes.add_textbox(self.left, top, self.width, heading_h)
                h_box.text_frame.word_wrap = True
                h_run = h_box.text_frame.paragraphs[0].add_run()
                h_run.text = heading
                _apply_text_style(h_run, self.heading_style)

            b_box = slide.shapes.add_textbox(self.left, top + heading_h + gap_after_heading, self.width, body_h)
            b_box.text_frame.word_wrap = True
            b_run = b_box.text_frame.paragraphs[0].add_run()
            b_run.text = body
            _apply_text_style(b_run, self.body_style)

        self._place(total_h, add)

    def add_table(self, label: str, columns: list[str], rows: list[list]):
        label_h = Inches(0.35) if label else 0
        row_h = Inches(0.4)
        table_h = row_h * (len(rows) + 1)
        total_h = label_h + table_h

        def add(slide, top):
            if label:
                l_box = slide.shapes.add_textbox(self.left, top, self.width, label_h)
                l_run = l_box.text_frame.paragraphs[0].add_run()
                l_run.text = label
                _apply_text_style(l_run, self.heading_style)
            table_top = top + label_h
            graphic_frame = slide.shapes.add_table(len(rows) + 1, len(columns), self.left, table_top, self.width, table_h)
            table = graphic_frame.table
            for c, col_label in enumerate(columns):
                table.cell(0, c).text = str(col_label)
            for r, row_values in enumerate(rows, start=1):
                for c, value in enumerate(row_values):
                    table.cell(r, c).text = str(value)
            _strip_table_style(table)

        self._place(total_h, add)


def render(
    company: str,
    service: str,
    style: str,
    field_values: dict,
    table_values: dict | None = None,
    section_values: dict | None = None,
    table_columns: dict | None = None,
    rep_signature_id: str | None = None,
    client_signature_id: str | None = None,
) -> Path:
    """Fills a shared pptx style template with client-submitted content plus the selected
    company's baked-in branding (name, rep contact, and signature image if one has been
    uploaded for that company). `service` only affects the output filename.

    Core sections, any extra sections, and any tables all flow together onto the same
    page(s) starting at `content_flow.first_slide` — stacked one under the next — and a
    new page is only added once the current one is actually full. Fixed-position fields
    (client name, contact block, branding, signature) are all applied FIRST, while slide
    numbers still match the schema exactly, before the flow engine starts inserting pages.

    Returns the path to the generated .pptx file.
    """
    table_values = table_values or {}
    table_columns = table_columns or {}
    section_values = section_values or {}
    schema = load_schema(style)
    companies = load_companies()
    company_data = companies[company]

    src = SHARED_DIR / style / "template.pptx"
    # A client name with a filesystem-reserved character (e.g. "Smith & Co: Phase 2",
    # "A/B Corp") would otherwise build an invalid path and crash generation outright --
    # strip anything that isn't a letter, digit, space, hyphen, or underscore.
    raw_client = field_values.get("client_name", "proposal").strip() or "proposal"
    safe_client = re.sub(r"[^\w\- ]+", "", raw_client).strip().replace(" ", "_") or "proposal"
    # An unreasonably long client name (copy-paste mistakes, pasted paragraphs) can push
    # the full output path past Windows' ~260 char limit, making file creation fail with
    # a raw FileNotFoundError -- cap it well before that.
    safe_client = safe_client[:60]
    # Two requests for the same company/service/style/client (a common case under any
    # concurrent load -- duplicate submissions, retries, multiple staff on the same
    # client) would otherwise collide on this exact filename and corrupt each other's
    # file mid-write; a short unique suffix keeps every generation's output isolated.
    out_name = f"{company}-{service}-{style}-{safe_client}-{uuid.uuid4().hex[:8]}.pptx"
    out_path = OUTPUT_DIR / out_name
    shutil.copy(src, out_path)

    prs = Presentation(out_path)
    badge_colors = schema.get("badge_colors")

    for removal in schema.get("remove_paragraphs", []):
        slide = prs.slides[removal["slide"] - 1]
        _remove_paragraphs(slide, removal["shape"], removal["paragraphs"])

    for field in schema.get("fields", []):
        key = field["key"]
        value = field_values.get(key, "")
        if value:
            for target in field.get("targets", []):
                slide = prs.slides[target["slide"] - 1]
                shape = _set_run_text(slide, target, value)
                # Only single-line-ish fields (names, dates, headings) get shrink-to-fit —
                # forcing a multi-line "paragraph" field onto one line would wreck it.
                if field.get("type") != "paragraph":
                    _enable_shrink_to_fit(slide, shape, prs.slide_width, prs.slide_height, badge_colors)
        elif field.get("removable"):
            for target in field.get("targets", []):
                slide = prs.slides[target["slide"] - 1]
                _try_remove_shape(slide, target["shape"])

    for brand_key, targets in schema.get("branding_targets", {}).items():
        lookup_key = "name" if brand_key == "company_name" else brand_key
        # A value typed into this specific generation's form (e.g. a different staff
        # member's name for this one proposal) always wins over the company's saved
        # default -- this tool is shared across multiple people/companies, so nothing
        # about who's actually sending should be permanently locked to the company.
        value = field_values.get(brand_key) or company_data.get(lookup_key, "")
        # companies.json leaves rep_date blank for every company (a fixed per-company date
        # wouldn't make sense) -- default to today's date instead of silently writing an
        # empty value after the "Date:" prefix.
        if not value and brand_key.endswith("_date"):
            value = datetime.now().strftime("%m/%d/%Y")
        for target in targets:
            slide = prs.slides[target["slide"] - 1]
            shape = _set_run_text(slide, target, value)
            _enable_shrink_to_fit(slide, shape, prs.slide_width, prs.slide_height, badge_colors)

    global_replacements = []
    for field in schema.get("find_replace_fields", []):
        value = field_values.get(field["key"], "")
        if value:
            for find in field["finds"]:
                global_replacements.append((find, value))
    for brand_key, finds in schema.get("branding_find_replace", {}).items():
        lookup_key = "name" if brand_key == "company_name" else brand_key
        value = field_values.get(brand_key) or company_data.get(lookup_key, "")
        if value:
            for find in finds:
                global_replacements.append((find, value))
    if global_replacements:
        _global_find_replace(prs, global_replacements, badge_colors)

    def _place_signature(slot_cfg, sig_file):
        if not slot_cfg or not sig_file:
            return
        slide = prs.slides[slot_cfg["slide"] - 1]
        image_source = io.BytesIO(sig_file) if isinstance(sig_file, bytes) else str(sig_file)
        slide.shapes.add_picture(
            image_source,
            Inches(slot_cfg["left_in"]),
            Inches(slot_cfg["top_in"]),
            width=Inches(slot_cfg["width_in"]),
            height=Inches(slot_cfg["height_in"]),
        )

    # Rep signature: whatever was uploaded for this specific generation, falling back to
    # the company's saved default (e.g. the usual person who signs) if none was given --
    # this tool is shared across multiple staff, so a different person can override it
    # per-proposal without needing to change the company's stored default.
    rep_sig_file = (temp_signature_path(rep_signature_id) if rep_signature_id else None) or signature_path(company)
    _place_signature(schema.get("rep_signature_slot") or schema.get("signature_slot"), rep_sig_file)

    # Client signature: only ever supplied per-generation (there's no per-company default
    # for a client, since the client changes with every proposal).
    client_sig_file = temp_signature_path(client_signature_id) if client_signature_id else None
    _place_signature(schema.get("client_signature_slot"), client_sig_file)

    # --- content flow: everything below happens LAST, since it inserts new slides and
    # shifts every slide index after the insertion point. Nothing above may reference a
    # slide number greater than content_flow.first_slide once this section runs.
    flow_cfg = schema.get("content_flow")
    if flow_cfg:
        for slide_num in sorted(flow_cfg.get("remove_slides", []), reverse=True):
            _remove_slide(prs, slide_num - 1)

        first_slide = prs.slides[flow_cfg["first_slide"] - 1]
        style_ref = flow_cfg["style_reference"]
        heading_style = _capture_text_style(first_slide, style_ref["heading_shape"], style_ref.get("heading_paragraph", 0))
        body_style = _capture_text_style(first_slide, style_ref["body_shape"], style_ref.get("body_paragraph", 0))
        if "heading_size_override_pt" in style_ref:
            heading_style["size"] = Pt(style_ref["heading_size_override_pt"])
        flow = ContentFlow(prs, flow_cfg, heading_style, body_style)

        for block_cfg in flow_cfg["core_blocks"]:
            if "fixed_heading" in block_cfg:
                heading = block_cfg["fixed_heading"]
            else:
                heading = field_values.get(block_cfg.get("heading_key", ""), "")
            body = field_values.get(block_cfg["body_key"], "")
            if heading or body:
                flow.add_text_block(heading, body)

        for item in section_values.get(flow_cfg.get("extra_sections_key"), []):
            heading = item.get("heading", "")
            body = item.get("body", "")
            if heading or body:
                flow.add_text_block(heading, body)

        for table_field in flow_cfg.get("tables", []):
            rows = table_values.get(table_field["key"])
            if rows:
                columns = table_columns.get(table_field["key"]) or table_field["columns"]
                flow.add_table(table_field.get("label", ""), columns, rows)

        flow.finish()

    # --- standalone table: for templates with no content_flow (letters, NDAs,
    # agreements) that just need one pricing table somewhere, without the full
    # flowing-page machinery. Placed either directly on an existing spare slide, or on a
    # freshly cloned one inserted at a specific point, per the schema's configuration.
    standalone_cfg = schema.get("standalone_table")
    if standalone_cfg:
        rows = table_values.get(standalone_cfg["key"])
        if rows:
            if "target_slide" in standalone_cfg:
                target_slide = prs.slides[standalone_cfg["target_slide"] - 1]
            elif "insert_blank_after_slide" in standalone_cfg:
                insert_at = standalone_cfg["insert_blank_after_slide"] - 1
                target_slide = insert_blank_slide(prs, insert_at, standalone_cfg.get("layout_index", -1))
            else:
                base_index = standalone_cfg["base_slide"] - 1
                insert_at = standalone_cfg["insert_after_slide"] - 1
                target_slide = duplicate_slide(prs, base_index, insert_at)

            columns = table_columns.get(standalone_cfg["key"]) or standalone_cfg["columns"]
            left = Inches(standalone_cfg["left_in"])
            top = Inches(standalone_cfg["top_in"])
            width = Inches(standalone_cfg["width_in"])

            label = standalone_cfg.get("label")
            if label:
                label_box = target_slide.shapes.add_textbox(left, top, width, Inches(0.4))
                label_run = label_box.text_frame.paragraphs[0].add_run()
                label_run.text = label
                label_run.font.bold = True
                label_run.font.size = Pt(18)
                top = top + Inches(0.45)

            n_rows = len(rows) + 1
            row_h = Inches(0.4)
            graphic_frame = target_slide.shapes.add_table(n_rows, len(columns), left, top, width, row_h * n_rows)
            table = graphic_frame.table
            for c, col_label in enumerate(columns):
                table.cell(0, c).text = str(col_label)
            for r, row_values in enumerate(rows, start=1):
                for c, value in enumerate(row_values):
                    table.cell(r, c).text = str(value)
            _strip_table_style(table)

    prs.save(out_path)
    return out_path


def _docx_set_run(run, size_pt=None, bold=None, color_rgb=None, font_name=None) -> None:
    from docx.shared import Pt as DocxPt
    from docx.shared import RGBColor as DocxRGB
    if size_pt is not None:
        run.font.size = DocxPt(size_pt)
    if bold is not None:
        run.font.bold = bold
    if color_rgb is not None:
        run.font.color.rgb = DocxRGB(*color_rgb)
    if font_name is not None:
        run.font.name = font_name


def _largest_picture_blob(slide):
    """Finds the biggest picture on a slide (almost always the full-bleed background art)
    and returns its raw image bytes, so the Word export can reuse the exact same cover/
    closing artwork as the pptx instead of a generic plain page."""
    best, best_area = None, 0
    for sh in slide.shapes:
        if sh.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
            area = (sh.width or 0) * (sh.height or 0)
            if area > best_area:
                best_area, best = area, sh
    return best.image.blob if best is not None else None


def _cover_text_overlays(slide) -> list[dict]:
    """Some styles' cover/closing slides aren't a single flat picture -- they have a live
    text box (e.g. the client's name) sitting on top of the background art, which is why
    that slide's picture alone doesn't carry it. render_docx/render_pdf_native only ever
    extracted the picture, so on those styles the client's name (or other cover text)
    silently never made it into the Word/PDF export's cover page. Called against the
    already-filled deck (not a blank template), so any text returned here is real,
    already-substituted content, not placeholder text.

    Only returns text shapes stacked *above* the largest picture in z-order. A shape
    listed *before* the picture in the XML is drawn first and then fully covered by that
    opaque picture in the actual design -- i.e. genuinely invisible there on purpose (seen
    with progressive-ux-agreement's closing slide, which has a leftover "YOU" text box
    sitting under its "Thank You" background art). Overlaying it anyway would add a
    duplicate, mismatched copy of text the design never actually shows.
    """
    shapes = list(slide.shapes)
    picture_indexes = [i for i, sh in enumerate(shapes) if sh.shape_type == 13]
    topmost_picture_index = max(picture_indexes) if picture_indexes else -1

    overlays = []
    for i, sh in enumerate(shapes):
        if sh.shape_type == 13 or not sh.has_text_frame or i < topmost_picture_index:
            continue
        # python-pptx joins multiple paragraphs/soft line-breaks with "\n"/"\x0b" -- drawn
        # literally via drawCentredString (a single-line call) that shows as a missing-glyph
        # box, not an actual line break, so collapse them to spaces for this one-line overlay.
        text = " ".join(sh.text_frame.text.replace("\x0b", "\n").split("\n")).strip()
        text = " ".join(text.split())
        if not text or not sh.width or not sh.height:
            continue
        run = None
        for para in sh.text_frame.paragraphs:
            if para.runs:
                run = para.runs[0]
                break
        size_pt = run.font.size.pt if run and run.font.size else 24
        bold = bool(run.font.bold) if run else False
        color_hex = None
        try:
            if run and run.font.color and run.font.color.type is not None:
                # An explicit RGB value resolves directly; a *theme* color (very common for
                # this kind of reversed/knockout text over a photo background -- e.g.
                # schemeClr "bg1") has no .rgb to read, so approximate from the scheme name:
                # background-slot colors read as white, text-slot colors as black. Wrong for
                # an unusual theme, but far closer than defaulting to black on every cover.
                if run.font.color.type == MSO_THEME_COLOR.NOT_THEME_COLOR:
                    color_hex = str(run.font.color.rgb)
                else:
                    theme_name = str(run.font.color.theme_color)
                    color_hex = "FFFFFF" if "BACKGROUND" in theme_name else "000000"
        except AttributeError:
            color_hex = None
        overlays.append({
            "text": text,
            "left_in": sh.left / 914400,
            "top_in": sh.top / 914400,
            "width_in": sh.width / 914400,
            "height_in": sh.height / 914400,
            "size_pt": size_pt,
            "bold": bold,
            "color_hex": color_hex,
        })
    return overlays


def render_docx(
    company: str,
    service: str,
    style: str,
    field_values: dict,
    table_values: dict | None = None,
    section_values: dict | None = None,
    table_columns: dict | None = None,
) -> Path:
    """Generates a styled, editable Word document carrying the same content as the pptx
    proposal -- headings, body text, and pricing tables -- with real formatting (colors,
    a consistent typeface, a shaded table header) rather than a bare content dump. It
    doesn't attempt to reproduce the slide design itself (gradients, backgrounds, badges).
    Driven by the same schema as render(), so it works across every template's field
    structure (content_flow-based or plain fields/find_replace_fields)."""
    import io
    import docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_SECTION
    from docx.shared import Pt as DocxPt, Inches as DocxInches, RGBColor as DocxRGB
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    # Each brand's Word export uses the same accent color as its own pptx template
    # (heading text color, table-header hex) instead of one generic look for every brand.
    DOCX_ACCENTS = {
        "grid-style": ((0x8A, 0x6D, 0x00), "F0C11E"),                 # Cali Web Studios: gold/black
        "letter-style": ((0xB8, 0x1D, 0x3A), "B81D3A"),                # American Web Builders: red/navy
        "progressive-ux-agreement": ((0x14, 0x1C, 0xE0), "141CE0"),    # Progressive UX: blue
        "nda-style": ((0xE8, 0x1A, 0x6E), "E81A6E"),                   # Web App Dev: pink/blue
        "grant-writing-style": ((0x6B, 0x8E, 0x23), "A9D153"),         # Grant Writing Inc: olive/lime
    }
    ACCENT, TABLE_HEADER_FILL = DOCX_ACCENTS.get(style, ((0xEA, 0x33, 0x23), "EA3323"))
    INK = (0x22, 0x22, 0x22)
    MUTED = (0x6B, 0x6B, 0x6B)

    table_values = table_values or {}
    table_columns = table_columns or {}
    section_values = section_values or {}
    schema = load_schema(style)
    companies = load_companies()
    company_data = companies[company]

    doc = docx.Document()

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = DocxPt(11)
    normal.font.color.rgb = DocxRGB(*INK)
    normal.paragraph_format.line_spacing = 1.2
    normal.paragraph_format.space_after = DocxPt(8)

    def set_margins(section, left=1, right=1, top=0.9, bottom=0.9):
        section.left_margin = DocxInches(left)
        section.right_margin = DocxInches(right)
        section.top_margin = DocxInches(top)
        section.bottom_margin = DocxInches(bottom)

    set_margins(doc.sections[0])

    # Reuse the pptx template's own cover/closing artwork so the Word export actually
    # looks like the branded proposal instead of a plain generic document -- full-bleed,
    # same as the slide, on its own zero-margin page.
    src_prs = Presentation(SHARED_DIR / style / "template.pptx")
    cover_blob = _largest_picture_blob(src_prs.slides[0])
    closing_blob = _largest_picture_blob(src_prs.slides[-1])
    page_w, page_h = doc.sections[0].page_width, doc.sections[0].page_height

    if cover_blob:
        set_margins(doc.sections[0], 0, 0, 0, 0)
        cover_p = doc.add_paragraph()
        cover_p.paragraph_format.space_after = DocxPt(0)
        cover_run = cover_p.add_run()
        cover_run.add_picture(io.BytesIO(cover_blob), width=page_w, height=page_h)
        content_section = doc.add_section(WD_SECTION.NEW_PAGE)
        set_margins(content_section)

    client_name = field_values.get("client_name", "")

    title_p = doc.add_paragraph()
    title_run = title_p.add_run(schema.get("label", "Proposal"))
    _docx_set_run(title_run, size_pt=28, bold=True, color_rgb=ACCENT, font_name="Calibri")
    title_p.paragraph_format.space_after = DocxPt(2)
    # A bottom border under the title reads as a real title rule, not just bold text.
    pPr = title_p._p.get_or_add_pPr()
    pPr.append(parse_xml(
        f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="18" w:space="4" w:color="{TABLE_HEADER_FILL}"/></w:pBdr>'
    ))

    if client_name:
        sub_p = doc.add_paragraph()
        sub_run = sub_p.add_run(f"Prepared for {client_name}")
        _docx_set_run(sub_run, size_pt=13, bold=True, color_rgb=INK)
    rep_p = doc.add_paragraph()
    rep_run = rep_p.add_run(f"Prepared by {company_data.get('name', '')}")
    _docx_set_run(rep_run, size_pt=10, color_rgb=MUTED)
    rep_p.paragraph_format.space_after = DocxPt(20)

    # The actual pptx keeps content-page headings and pricing tables plain dark text on
    # white -- brand color there is confined to the cover/closing art and the title rule.
    # Coloring every heading/table header (an earlier version of this Word export) reads
    # as a mismatch once you've seen the real slide, so those stay INK here to match.
    def add_heading(text, size_pt=17):
        p = doc.add_paragraph()
        run = p.add_run(text)
        _docx_set_run(run, size_pt=size_pt, bold=True, color_rgb=INK)
        p.paragraph_format.space_before = DocxPt(16)
        p.paragraph_format.space_after = DocxPt(6)
        return p

    def add_body(text):
        p = doc.add_paragraph(text)
        p.paragraph_format.space_after = DocxPt(10)
        return p

    def add_table(label, columns, rows):
        if label:
            add_heading(label, size_pt=15)
        table = doc.add_table(rows=1, cols=len(columns))
        table.style = "Table Grid"
        table.autofit = True
        header_cells = table.rows[0].cells
        for c, col_label in enumerate(columns):
            header_cells[c].text = ""
            p = header_cells[c].paragraphs[0]
            run = p.add_run(str(col_label))
            _docx_set_run(run, size_pt=11, bold=True, color_rgb=INK)
        for row_values in rows:
            cells = table.add_row().cells
            for c, value in enumerate(row_values):
                cells[c].text = ""
                run = cells[c].paragraphs[0].add_run(str(value))
                _docx_set_run(run, size_pt=10.5, color_rgb=INK)
        doc.add_paragraph().paragraph_format.space_after = DocxPt(4)

    flow_cfg = schema.get("content_flow")
    core_keys = set()
    if flow_cfg:
        for block_cfg in flow_cfg["core_blocks"]:
            if "heading_key" in block_cfg:
                core_keys.add(block_cfg["heading_key"])
            core_keys.add(block_cfg["body_key"])

        for block_cfg in flow_cfg["core_blocks"]:
            heading = block_cfg["fixed_heading"] if "fixed_heading" in block_cfg else field_values.get(block_cfg.get("heading_key", ""), "")
            body = field_values.get(block_cfg["body_key"], "")
            if heading or body:
                if heading:
                    add_heading(heading)
                if body:
                    add_body(body)

        for item in section_values.get(flow_cfg.get("extra_sections_key"), []):
            heading, body = item.get("heading", ""), item.get("body", "")
            if heading or body:
                if heading:
                    add_heading(heading)
                if body:
                    add_body(body)

        for table_field in flow_cfg.get("tables", []):
            rows = table_values.get(table_field["key"])
            if rows:
                columns = table_columns.get(table_field["key"]) or table_field["columns"]
                add_table(table_field.get("label", ""), columns, rows)

    # Any remaining text/paragraph field not already shown as part of a content_flow
    # block -- covers plain-fields templates (NDA-style, agreements) and fields like
    # contact info that sit alongside a content_flow template.
    all_fields = schema.get("fields", []) + schema.get("find_replace_fields", [])
    contact_lines = []
    for field in all_fields:
        key = field["key"]
        if key == "client_name" or key in core_keys:
            continue
        value = field_values.get(key, "")
        if not value:
            continue
        if field.get("type") == "paragraph":
            add_heading(field.get("label", key))
            add_body(value)
        else:
            contact_lines.append((field.get("label", key), value))

    if contact_lines:
        add_heading("Details", size_pt=15)
        for label, value in contact_lines:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = DocxPt(4)
            label_run = p.add_run(f"{label}:  ")
            _docx_set_run(label_run, bold=True, color_rgb=INK, size_pt=11)
            value_run = p.add_run(value)
            _docx_set_run(value_run, color_rgb=INK, size_pt=11)

    standalone_cfg = schema.get("standalone_table")
    if standalone_cfg:
        rows = table_values.get(standalone_cfg["key"])
        if rows:
            columns = table_columns.get(standalone_cfg["key"]) or standalone_cfg["columns"]
            add_table(standalone_cfg.get("label", ""), columns, rows)

    if closing_blob:
        closing_section = doc.add_section(WD_SECTION.NEW_PAGE)
        set_margins(closing_section, 0, 0, 0, 0)
        closing_p = doc.add_paragraph()
        closing_p.paragraph_format.space_after = DocxPt(0)
        closing_run = closing_p.add_run()
        closing_run.add_picture(io.BytesIO(closing_blob), width=page_w, height=page_h)

    raw_client = field_values.get("client_name", "proposal").strip() or "proposal"
    safe_client = re.sub(r"[^\w\- ]+", "", raw_client).strip().replace(" ", "_") or "proposal"
    safe_client = safe_client[:60]
    out_path = OUTPUT_DIR / f"{company}-{service}-{style}-{safe_client}-{uuid.uuid4().hex[:8]}.docx"
    doc.save(out_path)
    return out_path


def render_pdf_native(
    company: str,
    service: str,
    style: str,
    field_values: dict,
    table_values: dict | None = None,
    section_values: dict | None = None,
    table_columns: dict | None = None,
) -> Path:
    """Builds the proposal PDF directly with reportlab instead of converting the pptx via
    LibreOffice. LibreOffice isn't installable on most free/shared hosts (PythonAnywhere,
    plain shared cPanel, etc.), so this pure-Python path is what makes PDF export actually
    portable -- it trades exact slide-design fidelity for the same branded-cover-page +
    plain-content look already used by render_docx(), which is a real proposal, not a
    fallback dump. Mirrors render_docx()'s content-walking logic so both exports stay in
    sync as templates evolve."""
    import io
    from pypdf import PdfWriter, PdfReader
    from reportlab.pdfgen import canvas as pdfcanvas
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib.utils import ImageReader
    from reportlab.platypus import (
        Paragraph, Spacer, Table as RLTable, TableStyle,
        Image as RLImage, SimpleDocTemplate, HRFlowable,
    )

    DOCX_ACCENTS = {
        "grid-style": ((0x8A, 0x6D, 0x00), "F0C11E"),
        "letter-style": ((0xB8, 0x1D, 0x3A), "B81D3A"),
        "progressive-ux-agreement": ((0x14, 0x1C, 0xE0), "141CE0"),
        "nda-style": ((0xE8, 0x1A, 0x6E), "E81A6E"),
        "grant-writing-style": ((0x6B, 0x8E, 0x23), "A9D153"),
    }
    accent_rgb, _ = DOCX_ACCENTS.get(style, ((0xEA, 0x33, 0x23), "EA3323"))
    ACCENT = colors.Color(*[c / 255 for c in accent_rgb])
    ACCENT_TINT = colors.Color(*[(c + (255 - c) * 0.88) / 255 for c in accent_rgb])
    INK = colors.Color(0x22 / 255, 0x22 / 255, 0x22 / 255)
    MUTED = colors.Color(0x6B / 255, 0x6B / 255, 0x6B / 255)

    table_values = table_values or {}
    table_columns = table_columns or {}
    section_values = section_values or {}
    schema = load_schema(style)
    companies = load_companies()
    company_data = companies[company]

    # Some styles' cover/closing slides carry a live text box over the background art (the
    # client's name, most often) rather than having everything baked into one flat picture
    # -- extracting from a blank template would only ever get the picture and silently drop
    # that text, so this renders the real, already-filled deck first and reads the cover/
    # closing slides from that instead.
    filled_pptx_path = render(company, service, style, field_values, table_values, section_values, table_columns)
    try:
        filled_prs = Presentation(filled_pptx_path)
        cover_slide, closing_slide = filled_prs.slides[0], filled_prs.slides[-1]
        cover_blob = _largest_picture_blob(cover_slide)
        closing_blob = _largest_picture_blob(closing_slide)
        cover_overlays = _cover_text_overlays(cover_slide)
        closing_overlays = _cover_text_overlays(closing_slide)
    finally:
        filled_pptx_path.unlink(missing_ok=True)

    page_w, page_h = letter
    raw_client = field_values.get("client_name", "proposal").strip() or "proposal"
    safe_client = re.sub(r"[^\w\- ]+", "", raw_client).strip().replace(" ", "_") or "proposal"
    safe_client = safe_client[:60]
    out_path = OUTPUT_DIR / f"{company}-{service}-{style}-{safe_client}-{uuid.uuid4().hex[:8]}.pdf"

    # ParagraphStyle's `leading` (line height) does NOT auto-scale with `fontSize` -- it
    # defaults to a fixed small value unless given explicitly. Leaving it unset on a large
    # font (like the 26pt title below) understates how much vertical space that line
    # actually needs, so the next flowable starts drawing while the previous one's glyphs
    # are still visually extending into that space -- i.e. overlapping text. Every style
    # here sets leading to ~1.2x its font size to avoid that.
    styles = {
        "title": ParagraphStyle("title", fontName="Helvetica-Bold", fontSize=26, leading=31, textColor=ACCENT, spaceAfter=4, alignment=TA_LEFT),
        "subtitle": ParagraphStyle("subtitle", fontName="Helvetica-Bold", fontSize=13, leading=16, textColor=INK, spaceAfter=2),
        "byline": ParagraphStyle("byline", fontName="Helvetica", fontSize=10, leading=13, textColor=MUTED, spaceAfter=18),
        "heading": ParagraphStyle("heading", fontName="Helvetica-Bold", fontSize=15, leading=18, textColor=INK, spaceBefore=14, spaceAfter=6),
        "body": ParagraphStyle("body", fontName="Helvetica", fontSize=10.5, textColor=INK, leading=15, spaceAfter=10),
        "label": ParagraphStyle("label", fontName="Helvetica-Bold", fontSize=10.5, leading=13, textColor=INK, spaceAfter=4),
    }

    def full_bleed_pdf_page(blob, overlays=None) -> bytes:
        """A single full-bleed image page, built with a raw canvas (no Platypus frames
        involved at all) -- kept as its own tiny PDF and merged in with pypdf afterwards.
        Mixing a zero-margin full-bleed page and a normally-margined content flow in one
        BaseDocTemplate via NextPageTemplate/PageTemplate switching turned out fragile (it
        produced a page where the title got drawn twice, overlapping, at the wrong scale) --
        building each part as an independent, simple PDF and merging them sidesteps that
        entirely."""
        buf = io.BytesIO()
        c = pdfcanvas.Canvas(buf, pagesize=letter)
        c.drawImage(ImageReader(io.BytesIO(blob)), 0, 0, width=page_w, height=page_h)
        for ov in (overlays or []):
            font = "Helvetica-Bold" if ov["bold"] else "Helvetica"
            c.setFont(font, ov["size_pt"])
            c.setFillColor(colors.HexColor(f"#{ov['color_hex']}") if ov["color_hex"] else colors.black)
            cx = (ov["left_in"] + ov["width_in"] / 2) * inch
            cy = page_h - (ov["top_in"] + ov["height_in"] / 2) * inch - ov["size_pt"] * 0.18
            c.drawCentredString(cx, cy, ov["text"])
        c.showPage()
        c.save()
        return buf.getvalue()

    story = []
    client_name = field_values.get("client_name", "")
    story.append(Paragraph(schema.get("label", "Proposal"), styles["title"]))
    if client_name:
        story.append(Paragraph(f"Prepared for {client_name}", styles["subtitle"]))
    story.append(Paragraph(f"Prepared by {company_data.get('name', '')}", styles["byline"]))
    story.append(HRFlowable(width="100%", thickness=2, color=ACCENT, spaceAfter=16, spaceBefore=0))

    def add_heading(text):
        story.append(Paragraph(text, styles["heading"]))
        story.append(HRFlowable(width="100%", thickness=0.75, color=ACCENT_TINT, spaceAfter=8, spaceBefore=0))

    def add_body(text):
        story.append(Paragraph(text.replace("\n", "<br/>"), styles["body"]))

    def add_table(label, columns, rows):
        if label:
            add_heading(label)
        data = [columns] + [[str(v) for v in row] for row in rows]
        t = RLTable(data, hAlign="LEFT", colWidths=(page_w - 1.4 * inch) / max(len(columns), 1))
        t.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("TEXTCOLOR", (0, 0), (-1, 0), ACCENT),
            ("TEXTCOLOR", (0, 1), (-1, -1), INK),
            ("BACKGROUND", (0, 0), (-1, 0), ACCENT_TINT),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FAFAFA")]),
            ("GRID", (0, 0), (-1, -1), 0.75, colors.HexColor("#DDDDDD")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.append(t)
        story.append(Spacer(1, 12))

    flow_cfg = schema.get("content_flow")
    core_keys = set()
    if flow_cfg:
        for block_cfg in flow_cfg["core_blocks"]:
            if "heading_key" in block_cfg:
                core_keys.add(block_cfg["heading_key"])
            core_keys.add(block_cfg["body_key"])

        for block_cfg in flow_cfg["core_blocks"]:
            heading = block_cfg["fixed_heading"] if "fixed_heading" in block_cfg else field_values.get(block_cfg.get("heading_key", ""), "")
            body = field_values.get(block_cfg["body_key"], "")
            if heading or body:
                if heading:
                    add_heading(heading)
                if body:
                    add_body(body)

        for item in section_values.get(flow_cfg.get("extra_sections_key"), []):
            heading, body = item.get("heading", ""), item.get("body", "")
            if heading or body:
                if heading:
                    add_heading(heading)
                if body:
                    add_body(body)

        for table_field in flow_cfg.get("tables", []):
            rows = table_values.get(table_field["key"])
            if rows:
                columns = table_columns.get(table_field["key"]) or table_field["columns"]
                add_table(table_field.get("label", ""), columns, rows)

    all_fields = schema.get("fields", []) + schema.get("find_replace_fields", [])
    contact_lines = []
    for field in all_fields:
        key = field["key"]
        if key == "client_name" or key in core_keys:
            continue
        value = field_values.get(key, "")
        if not value:
            continue
        if field.get("type") == "paragraph":
            add_heading(field.get("label", key))
            add_body(value)
        else:
            contact_lines.append((field.get("label", key), value))

    if contact_lines:
        add_heading("Details")
        for label, value in contact_lines:
            story.append(Paragraph(f"<b>{label}:</b> {value}", styles["body"]))

    standalone_cfg = schema.get("standalone_table")
    if standalone_cfg:
        rows = table_values.get(standalone_cfg["key"])
        if rows:
            columns = table_columns.get(standalone_cfg["key"]) or standalone_cfg["columns"]
            add_table(standalone_cfg.get("label", ""), columns, rows)

    def draw_footer(c, doc_):
        c.saveState()
        c.setStrokeColor(colors.HexColor("#DDDDDD"))
        c.setLineWidth(0.5)
        c.line(margin, 0.5 * inch, page_w - margin, 0.5 * inch)
        c.setFont("Helvetica", 8)
        c.setFillColor(MUTED)
        c.drawString(margin, 0.34 * inch, company_data.get("name", ""))
        c.drawRightString(page_w - margin, 0.34 * inch, f"Page {doc_.page}")
        c.restoreState()

    content_buf = io.BytesIO()
    margin = 0.7 * inch
    content_doc = SimpleDocTemplate(
        content_buf, pagesize=letter,
        leftMargin=margin, rightMargin=margin, topMargin=margin, bottomMargin=margin,
    )
    content_doc.build(story, onFirstPage=draw_footer, onLaterPages=draw_footer)
    content_buf.seek(0)

    writer = PdfWriter()
    if cover_blob:
        writer.append(PdfReader(io.BytesIO(full_bleed_pdf_page(cover_blob, cover_overlays))))
    writer.append(PdfReader(content_buf))
    if closing_blob:
        writer.append(PdfReader(io.BytesIO(full_bleed_pdf_page(closing_blob, closing_overlays))))
    with open(out_path, "wb") as f:
        writer.write(f)
    return out_path


_LIBREOFFICE_FALLBACK_PATHS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]

# Each LibreOffice conversion below gets its own throwaway profile so concurrent calls
# don't corrupt a shared one, but each headless instance is still a full, heavy LibreOffice
# process -- launching too many at once starves the machine's CPU/memory and individual
# launches start failing outright (non-zero exit) rather than just queueing. Capping how
# many run at once keeps every conversion isolated *and* reliable under real concurrent load.
_LIBREOFFICE_CONCURRENCY = threading.BoundedSemaphore(1)


def render_pdf(pptx_path: Path) -> Path | None:
    """Converts a pptx to pdf via headless LibreOffice, if available. Returns None if soffice isn't installed."""
    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if not soffice:
        # A fresh LibreOffice install doesn't land on PATH until the shell/process
        # restarts -- fall back to the standard install location so PDF export works
        # immediately without requiring a server restart.
        soffice = next((p for p in _LIBREOFFICE_FALLBACK_PATHS if Path(p).exists()), None)
    if not soffice:
        return None
    # Even with an isolated profile and bounded concurrency, headless LibreOffice still
    # occasionally exits non-zero for no reproducible reason (a transient flake under load,
    # not a bad input -- the same file converts fine on retry) -- one retry with a brand new
    # profile clears these without surfacing a spurious failure to the user.
    last_error: subprocess.CalledProcessError | None = None
    for attempt in range(3):
        with tempfile.TemporaryDirectory(prefix="qp_soffice_") as profile_dir, _LIBREOFFICE_CONCURRENCY:
            profile_uri = Path(profile_dir).resolve().as_uri()
            try:
                subprocess.run(
                    [
                        soffice, "--headless", "--norestore",
                        f"-env:UserInstallation={profile_uri}",
                        "--convert-to", "pdf", "--outdir", str(OUTPUT_DIR), str(pptx_path),
                    ],
                    check=True,
                    timeout=60,
                )
                return pptx_path.with_suffix(".pdf")
            except subprocess.CalledProcessError as e:
                last_error = e
    raise last_error
