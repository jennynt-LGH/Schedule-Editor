#!/usr/bin/env python3
"""
pdf_text_editor.py

Find-and-replace / delete text inside a PDF, driven by an Excel
instructions file, while trying to keep the new text matching the
original font style, size, and position.

Instructions file format (first sheet):

    Table 1 (a row containing the header "Replace from this" starts it):
        Replace from this | to this | Exception | Font size of new text | (anything else, ignored)
        <old text>         | <new>   | page 9    | 12.08
        ...
        (a blank row ends the table)

    Table 2 (a row containing "Delete these words from PDF" starts it):
        <word or phrase to delete>
        <word or phrase to delete>
        ...
        (a blank row ends the table)

"Exception" may be blank, "-", or something like "page 9" / "pages 3, 5"
-- any numbers found in that cell are treated as 1-indexed page numbers
to skip for that rule.

Anything that is NOT part of the two tables above is ignored by design,
so the spreadsheet can carry human-readable notes without confusing the
parser. Concretely:
    - Any row(s) above the "Replace from this" header row (e.g. a title,
      or a "Guide for users" row explaining how to fill the sheet in) are
      skipped, since scanning only starts once that exact header is found.
    - Each table ends at its first fully blank row. Anything below that
      blank row -- e.g. a closing reminder like "Before finalising, add a
      visual check of the whole page..." -- is never read as data, even
      if it's in the same column as the delete-phrase list above it.
This means notes/instructions meant for a *person* filling in the sheet
(or for whoever reviews the finished PDF) can sit right in the sheet
without needing to be removed before uploading it.

Usage:
    python pdf_text_editor.py INPUT.pdf INSTRUCTIONS.xlsx OUTPUT.pdf [--previews DIR]
"""

import argparse
import os
import re
import sys
import tempfile

import fitz  # PyMuPDF
import openpyxl


# ---------------------------------------------------------------------------
# Instructions parsing
# ---------------------------------------------------------------------------

def _norm(cell):
    return str(cell).strip().lower() if cell is not None else ""


def parse_instructions(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))

    replace_rules = []
    delete_phrases = []

    i, n = 0, len(rows)
    while i < n:
        row = rows[i]
        norm_row = [_norm(c) for c in row]

        if any("replace from this" in c for c in norm_row):
            col_from = next(j for j, c in enumerate(norm_row) if "replace from this" in c)
            col_to = next((j for j, c in enumerate(norm_row) if "to this" in c), col_from + 1)
            col_exc = next((j for j, c in enumerate(norm_row) if "exception" in c), col_from + 2)
            col_size = next((j for j, c in enumerate(norm_row) if "font size" in c), col_from + 3)
            i += 1
            while i < n:
                r = rows[i]
                old = r[col_from] if col_from < len(r) else None
                if old is None or str(old).strip() == "":
                    break
                new = r[col_to] if col_to < len(r) else None
                exc_raw = r[col_exc] if col_exc < len(r) else None
                size_raw = r[col_size] if col_size < len(r) else None

                exception_pages = set()
                if exc_raw is not None and str(exc_raw).strip() not in ("", "-"):
                    exception_pages = {int(x) for x in re.findall(r"\d+", str(exc_raw))}

                size = None
                try:
                    if size_raw not in (None, ""):
                        size = float(size_raw)
                except (TypeError, ValueError):
                    size = None

                replace_rules.append({
                    "old": str(old).strip(),
                    "new": "" if new is None else str(new).strip(),
                    "exception_pages": exception_pages,
                    "size": size,
                })
                i += 1
            continue

        if any("delete these words" in c for c in norm_row):
            i += 1
            while i < n:
                r = rows[i]
                vals = [c for c in r if c is not None and str(c).strip() != ""]
                if not vals:
                    break
                delete_phrases.append(str(vals[0]).strip())
                i += 1
            continue

        i += 1

    return replace_rules, delete_phrases


# ---------------------------------------------------------------------------
# PDF helpers
# ---------------------------------------------------------------------------

def get_spans(page):
    spans = []
    d = page.get_text("dict")
    for block in d["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                spans.append(span)
    return spans


def find_containing_span(spans, rect):
    cy = (rect.y0 + rect.y1) / 2
    for span in spans:
        bx0, by0, bx1, by1 = span["bbox"]
        if by0 - 1 <= cy <= by1 + 1 and bx0 - 1 <= rect.x0 and rect.x1 <= bx1 + 1:
            return span
    return None


def extract_all_fonts(doc, workdir):
    """Extract every embedded font in the document to disk, keyed by base
    font name (subset prefix like 'ABCDEF+' stripped)."""
    font_files = {}
    seen = set()
    for page in doc:
        for f in page.get_fonts(full=True):
            xref = f[0]
            if xref in seen:
                continue
            seen.add(xref)
            base_name = f[3]
            if "+" in base_name:
                base_name = base_name.split("+", 1)[1]
            try:
                info = doc.extract_font(xref)
            except Exception:
                continue
            ext, buf = info[1], info[3]
            if not buf:
                continue
            path = os.path.join(workdir, f"{base_name}.{ext or 'ttf'}")
            with open(path, "wb") as fh:
                fh.write(buf)
            font_files[base_name] = path
    return font_files


def find_overlap_warnings(page, inserted_specs, cover_rects):
    """Automated version of "add a visual check of the whole page (not
    just the edited spot) to catch any unintended text overlaps".

    inserted_specs: list of (x, y, text) tuples for text we just inserted
    on this page.
    cover_rects: the white-out boxes drawn on this page -- used to ignore
    the old text we deliberately hid underneath them (that text is still
    technically present/extractable, by design -- see the module
    docstring -- so it would otherwise "overlap" its own replacement on
    every single edit and drown out real warnings).

    Re-reads the page's text after insertion and flags any case where one
    of *our* inserted spans overlaps a bounding box of some other,
    genuinely still-visible span on the page -- the exact failure pattern
    that previously caused things like a stray leftover "-" bleeding into
    replacement text. Returns a list of human-readable warning strings
    (empty if nothing looks wrong).
    """
    spans = get_spans(page)
    inserted_spans, other_spans = [], []
    for span in spans:
        ox, oy = span["origin"]
        is_ours = any(
            abs(ox - ix) < 1.0 and abs(oy - iy) < 1.0
            for ix, iy, _ in inserted_specs
        )
        (inserted_spans if is_ours else other_spans).append(span)

    warnings = []
    for ins in inserted_spans:
        ins_rect = fitz.Rect(ins["bbox"])
        for other in other_spans:
            other_rect = fitz.Rect(other["bbox"])
            inter = ins_rect & other_rect
            if inter.is_empty:
                continue
            inter_area = inter.width * inter.height
            if inter_area < 1.0:
                continue
            # The old/other text can perfectly legitimately share space
            # with our new text -- that's exactly what happens when we
            # white-out old text and write new text in its place. Only
            # the OVERLAPPING REGION itself needs to be painted over for
            # this to be invisible in the final render; the rest of that
            # other span (e.g. an untouched "Handle:" label before it)
            # is irrelevant. So check coverage of the intersection, not
            # of the whole other span.
            covered = any(
                (inter & cover).width * (inter & cover).height >= 0.9 * inter_area
                for cover in cover_rects
            )
            if covered:
                continue
            warnings.append(
                f"'{ins['text'].strip()}' may overlap "
                f"'{other['text'].strip()}' -- check this page closely"
            )
    return warnings


def resolve_font(span_font_name, font_files):
    """Return (fontname, fontfile_or_None) for insert_text()."""
    name = span_font_name or ""
    if "+" in name:
        name = name.split("+", 1)[1]
    if name in font_files:
        return name, font_files[name]
    lname = name.lower()
    if "bold" in lname and ("italic" in lname or "oblique" in lname):
        return "hebi", None
    if "bold" in lname:
        return "hebo", None
    if "italic" in lname or "oblique" in lname:
        return "heit", None
    return "helv", None


# ---------------------------------------------------------------------------
# Main processing
# ---------------------------------------------------------------------------

def process(input_pdf, xlsx_path, output_pdf, preview_dir=None):
    replace_rules, delete_phrases = parse_instructions(xlsx_path)
    doc = fitz.open(input_pdf)
    workdir = tempfile.mkdtemp()
    font_files = extract_all_fonts(doc, workdir)

    total_replaced = 0
    total_deleted = 0
    modified_pages = set()
    overlap_warnings = {}  # page_num -> list of warning strings

    for pno in range(len(doc)):
        page = doc[pno]
        page_num = pno + 1
        spans = get_spans(page)

        cover_rects = []
        insert_jobs = []  # (x, y, text, fontname, fontfile_or_None, size)

        for rule in replace_rules:
            old, new = rule["old"], rule["new"]
            if not old or page_num in rule["exception_pages"]:
                continue
            for rect in page.search_for(old):
                span = find_containing_span(spans, rect)
                pad = 0.4
                if span is None:
                    print(f"  [warn] page {page_num}: could not find styling for "
                          f"'{old}' -- using a fallback font/position", file=sys.stderr)
                    baseline_y = rect.y1 - (rect.y1 - rect.y0) * 0.2
                    fontkey, fontfile = "helv", None
                    size = rule["size"] or (rect.y1 - rect.y0) * 0.8
                    span_x1 = rect.x1
                    suffix = ""
                else:
                    baseline_y = span["origin"][1]
                    fontkey, fontfile = resolve_font(span["font"], font_files)
                    size = rule["size"] if rule["size"] else span["size"]
                    span_x1 = span["bbox"][2]
                    idx = span["text"].find(old)
                    suffix = span["text"][idx + len(old):] if idx != -1 else ""

                # Cover from the start of the old text through to the end of
                # its span (so any trailing text sharing the same span, e.g.
                # "OTHER - 3 quantity:", gets reflowed after the new text
                # instead of being overlapped by it).
                cover_rects.append(fitz.Rect(rect.x0 - pad, rect.y0 - pad, span_x1 + pad, rect.y1 + pad))
                insert_jobs.append((rect.x0, baseline_y, new + suffix, fontkey, fontfile, size))
                total_replaced += 1
                modified_pages.add(page_num)

        for phrase in delete_phrases:
            for rect in page.search_for(phrase):
                span = find_containing_span(spans, rect)
                pad = 0.4
                if span is not None:
                    remainder = span["text"].replace(phrase, "", 1)
                    if remainder.strip(" \t-:") == "":
                        # The whole span is essentially just this phrase
                        # (plus separators like " - ") -- remove all of it
                        # so no dangling punctuation is left behind.
                        bx0, by0, bx1, by1 = span["bbox"]
                        cover_rects.append(fitz.Rect(bx0 - pad, by0 - pad, bx1 + pad, by1 + pad))
                        total_deleted += 1
                        modified_pages.add(page_num)
                        continue
                cover_rects.append(fitz.Rect(rect.x0 - pad, rect.y0 - pad, rect.x1 + pad, rect.y1 + pad))
                total_deleted += 1
                modified_pages.add(page_num)

        if not cover_rects:
            continue

        # Non-destructive white overlay (rather than true redaction) --
        # redacting text can corrupt shared embedded font glyphs elsewhere
        # on the same page. The tradeoff: old text is visually hidden but
        # technically still present/extractable underneath.
        for r in cover_rects:
            page.draw_rect(r, color=None, fill=(1, 1, 1), fill_opacity=1, overlay=True)

        for x, y, text, fontkey, fontfile, size in insert_jobs:
            if fontfile:
                page.insert_text((x, y), text, fontsize=size, fontname=fontkey,
                                  fontfile=fontfile, color=(0, 0, 0))
            else:
                page.insert_text((x, y), text, fontsize=size, fontname=fontkey, color=(0, 0, 0))

        # Automated stand-in for a human eyeballing "the whole page, not
        # just the edited spot": re-check the page we just edited for any
        # inserted text unexpectedly overlapping something else.
        if insert_jobs:
            inserted_specs = [(x, y, text) for x, y, text, *_ in insert_jobs]
            page_warnings = find_overlap_warnings(page, inserted_specs, cover_rects)
            if page_warnings:
                overlap_warnings[page_num] = page_warnings
                for w in page_warnings:
                    print(f"  [OVERLAP WARNING] page {page_num}: {w}", file=sys.stderr)

    doc.save(output_pdf)

    if preview_dir and modified_pages:
        os.makedirs(preview_dir, exist_ok=True)
        preview_doc = fitz.open(output_pdf)
        for pno in sorted(modified_pages):
            preview_doc[pno - 1].get_pixmap(dpi=150).save(
                os.path.join(preview_dir, f"page{pno}_preview.png"))

    return total_replaced, total_deleted, sorted(modified_pages), overlap_warnings


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("pdf", help="Input PDF file")
    parser.add_argument("xlsx", help="Instructions .xlsx file")
    parser.add_argument("output", help="Output PDF path")
    parser.add_argument("--previews", default=None, help="Directory to save preview PNGs of modified pages")
    args = parser.parse_args()

    replaced, deleted, pages, overlap_warnings = process(args.pdf, args.xlsx, args.output, args.previews)
    print(f"Replaced {replaced} instance(s), deleted {deleted} instance(s).")
    print(f"Modified pages: {pages}")
    if overlap_warnings:
        print("\n*** POSSIBLE TEXT OVERLAP DETECTED -- review these pages before use: ***")
        for pno, warnings in overlap_warnings.items():
            print(f"  Page {pno}:")
            for w in warnings:
                print(f"    - {w}")
    else:
        print("No overlap issues detected on the modified pages.")
