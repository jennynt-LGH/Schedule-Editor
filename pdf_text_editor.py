#!/usr/bin/env python3
"""
pdf_text_editor.py

Find-and-replace / delete text inside a PDF, driven by an Excel
instructions file, while trying to keep the new text matching the
original font style, size, and position.

Instructions file format (first sheet):

    Table 1 (a row containing the header "Replace from this" starts it):
        Replace from this | to this | Only here | Skip this (don't apply to these) | Font size of new text | (anything else, ignored)
        <old text>         | <new>   | Pos no. 12 | page 9          | 12.08
        ...
        (a blank row ends the table)

    Table 2 (a row containing "Delete these words from PDF" starts it),
    which optionally takes the SAME "Only here" / "Skip this" columns:
        Delete these words from PDF | Only here | Skip this (don't apply to these)
        <word or phrase to delete>   | Pos no. 3  |
        the whole line of "<text>"   |            | page 9
        ...
        (a blank row ends the table)

A delete row's phrase can either be:
    - a literal word/phrase that appears in the PDF -- only that text is
      removed, or
    - "the whole line of "<text>"" -- finds <text> in the PDF, then
      removes the ENTIRE line it's on (useful for removing a whole
      "U-value (W/m2K)= 1.39" style line by only naming part of it).

"Only here" and "Skip this (don't apply to these)" both accept the same
kind of value in either table: blank, "-", a page number/numbers ("page
9", "pages 3, 5"), or a POS number/numbers ("Pos no. 12", "POS #3, 7").
Whether a cell means pages or POS numbers is decided by whether the word
"pos" appears in it anywhere (any spacing/punctuation/case: "POS 12",
"pos.no 12", "POS #12" all count) -- otherwise the numbers in it are
treated as page numbers.
    - "Only here": the rule applies ONLY at the listed page(s)/POS
      number(s), and is skipped everywhere else in the document.
    - "Skip this (don't apply to these)": the rule applies everywhere
      EXCEPT the listed page(s)/POS number(s) (this column has also been
      called "Exception" / "Except for this" in older sheets -- all three
      headers are recognized).
A POS number refers to the "Pos.no N:" label the PDF itself prints next
to each item -- not a spreadsheet row number.

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


def _expand_range_tokens(s):
    """Turn a comma-separated cell of numbers/ranges into a set of ints.

    Each comma-separated chunk is either a single number ("9") or a range
    ("3-7"). A chunk counts as a range whenever it contains a "-" AND at
    least two numbers -- so "POS 1-8", "1-3", "POS 1-POS 6", and "6-POS 9"
    all resolve the same way (only the digits and the dash matter; any
    words like "POS" in between are ignored). A lone number, or a chunk
    with a dash but only one number in it (e.g. a stray trailing "-"), is
    treated as a single value rather than a range. Whichever of the two
    numbers is larger becomes the range's end, so "8-3" behaves the same
    as "3-8".
    """
    nums = set()
    for chunk in s.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        parts = re.findall(r"\d+", chunk)
        if not parts:
            continue
        if "-" in chunk and len(parts) >= 2:
            start, end = int(parts[0]), int(parts[-1])
            if start > end:
                start, end = end, start
            nums.update(range(start, end + 1))
        else:
            nums.update(int(p) for p in parts)
    return nums


def _parse_location_cell(raw):
    """Parse an "Only here" / "Except for this" style cell. Returns
    (pages, pos_numbers) -- a pair of sets of ints, at most one non-empty.

    Accepts blank, "-", a page reference ("page 9", "pages 3, 5",
    "pages 3-7", "page 1-page 6"), or a POS-number reference ("Pos no.
    12", "POS #3, 7", "POS 1-8", "POS 1-3, POS 6-8", "POS 1-POS 6").
    Ranges ("a-b") are expanded to every number in between, inclusive.
    Whether a cell means pages or POS numbers is decided by whether "pos"
    appears anywhere in the cell (case-insensitive, any spacing/
    punctuation); otherwise any numbers found are treated as page
    numbers. Matching is case-insensitive throughout.
    """
    if raw is None:
        return set(), set()
    s = str(raw).strip()
    if s == "" or s == "-":
        return set(), set()
    nums = _expand_range_tokens(s)
    if not nums:
        return set(), set()
    if "pos" in s.lower():
        return set(), nums
    return nums, set()


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
            # "Only here" is the newer inclusion column; older sheets won't
            # have it at all, so col_only can legitimately stay None.
            col_only = next((j for j, c in enumerate(norm_row) if "only" in c), None)
            # "except"/"skip" covers all the header wordings seen so far:
            # the original "Exception", and the newer "Except for this" /
            # "Skip this (don't apply to these)".
            col_exc = next(
                (j for j, c in enumerate(norm_row) if "except" in c or "skip" in c),
                (col_only + 1) if col_only is not None else col_from + 2,
            )
            col_size = next((j for j, c in enumerate(norm_row) if "font size" in c), col_exc + 1)
            i += 1
            while i < n:
                r = rows[i]
                old = r[col_from] if col_from < len(r) else None
                if old is None or str(old).strip() == "":
                    break
                new = r[col_to] if col_to < len(r) else None
                only_raw = r[col_only] if (col_only is not None and col_only < len(r)) else None
                exc_raw = r[col_exc] if col_exc < len(r) else None
                size_raw = r[col_size] if col_size < len(r) else None

                only_pages, only_pos = _parse_location_cell(only_raw)
                except_pages, except_pos = _parse_location_cell(exc_raw)

                size = None
                try:
                    if size_raw not in (None, ""):
                        size = float(size_raw)
                except (TypeError, ValueError):
                    size = None

                replace_rules.append({
                    "old": str(old).strip(),
                    "new": "" if new is None else str(new).strip(),
                    "only_pages": only_pages,
                    "only_pos": only_pos,
                    "except_pages": except_pages,
                    "except_pos": except_pos,
                    "size": size,
                })
                i += 1
            continue

        if any("delete these words" in c for c in norm_row):
            col_phrase = next(j for j, c in enumerate(norm_row) if "delete these words" in c)
            col_only = next((j for j, c in enumerate(norm_row) if "only" in c), None)
            col_skip = next((j for j, c in enumerate(norm_row) if "except" in c or "skip" in c), None)
            i += 1
            while i < n:
                r = rows[i]
                phrase = r[col_phrase] if col_phrase < len(r) else None
                if phrase is None or str(phrase).strip() == "":
                    break
                only_raw = r[col_only] if (col_only is not None and col_only < len(r)) else None
                skip_raw = r[col_skip] if (col_skip is not None and col_skip < len(r)) else None
                only_pages, only_pos = _parse_location_cell(only_raw)
                except_pages, except_pos = _parse_location_cell(skip_raw)
                delete_phrases.append({
                    "phrase": str(phrase).strip(),
                    "only_pages": only_pages,
                    "only_pos": only_pos,
                    "except_pages": except_pages,
                    "except_pos": except_pos,
                })
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
            line_bbox = line["bbox"]
            for span in line["spans"]:
                span = dict(span)
                span["line_bbox"] = line_bbox
                spans.append(span)
    return spans


_POS_LABEL_RE = re.compile(r"pos\.?\s*no\.?\s*(\d+)", re.IGNORECASE)


def build_pos_blocks(spans, page_bottom):
    """Each item in the schedule PDF is introduced by a "Pos.no N:" label.
    This finds every such label on the page and works out the vertical
    range of the page that belongs to each one (from its own label down
    to the start of the next one, or the bottom of the page for the last
    item) -- so a match's y-position can be mapped back to "which POS
    number is this text part of". Returns a list of (y_start, y_end,
    pos_number) tuples.
    """
    labels = []
    for span in spans:
        m = _POS_LABEL_RE.search(span["text"])
        if m:
            labels.append((span["bbox"][1], int(m.group(1))))
    labels.sort(key=lambda t: t[0])
    blocks = []
    for idx, (y0, pos_num) in enumerate(labels):
        y_end = labels[idx + 1][0] if idx + 1 < len(labels) else page_bottom
        blocks.append((y0, y_end, pos_num))
    return blocks


def pos_number_for_rect(pos_blocks, rect):
    """Which POS number's block a given match rectangle falls into, or
    None if the page has no recognizable "Pos.no N:" labels at all."""
    cy = (rect.y0 + rect.y1) / 2
    for y_start, y_end, pos_num in pos_blocks:
        if y_start <= cy < y_end:
            return pos_num
    return None


_WHOLE_LINE_RE = re.compile(r"whole line(?:s)?\s+of\b\s*(.*)", re.IGNORECASE)


def parse_delete_phrase(raw):
    """A "delete" row is normally a literal phrase to remove. But it can
    also be written as an instruction FOR A PERSON describing what to
    remove, e.g. 'the whole line of "U-value (W/m2K)"' -- that sentence
    itself will never appear in the PDF, so searching for it literally
    would silently find and delete nothing.

    Detect that pattern and return (search_text, whole_line) where
    search_text is the actual text to locate on the page, and whole_line
    means: once found, remove the ENTIRE line it's on (not just the
    matched substring). Anything that doesn't match the pattern is
    treated as a plain literal phrase, unchanged from before.
    """
    s = raw.strip()
    m = _WHOLE_LINE_RE.search(s)
    if m:
        inner = m.group(1).strip()
        inner = inner.strip("\"'“”‘’").strip()
        inner = inner.rstrip(".").strip()
        if inner:
            return inner, True
    return s, False


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


def _same_text_exact_case(extracted, expected):
    """Case-sensitive equality that tolerates whitespace differences.
    PyMuPDF's get_textbox() reconstructs spacing from glyph positions and
    can occasionally introduce or drop a space compared to the literal
    string a rule was typed with -- e.g. a slightly different inter-word
    gap in one occurrence of an otherwise-identical phrase can make an
    exact string comparison fail even though the actual letters (and
    their case) match perfectly. Collapse runs of whitespace to a single
    space on both sides before comparing so only real content differences
    -- including case -- cause a mismatch."""
    return " ".join(extracted.split()) == " ".join(expected.split())


def _base14_for_style(name):
    """Pick a built-in base-14 font (full standard glyph coverage) that
    matches the bold/italic style implied by an original font's name."""
    lname = (name or "").lower()
    if "bold" in lname and ("italic" in lname or "oblique" in lname):
        return "hebi"
    if "bold" in lname:
        return "hebo"
    if "italic" in lname or "oblique" in lname:
        return "heit"
    return "helv"


def resolve_font(span_font_name, font_files):
    """Return (fontname, fontfile_or_None) for insert_text()."""
    name = span_font_name or ""
    if "+" in name:
        name = name.split("+", 1)[1]
    if name in font_files:
        return name, font_files[name]
    return _base14_for_style(name), None


def font_covers_text(fontkey, fontfile, text):
    """PDFs typically only embed the exact glyphs the original document
    used. A replacement/inserted phrase can easily need a character the
    original never did (e.g. a digit, an accent, a punctuation mark) --
    inserting with a font missing that glyph renders as a blank box
    ("tofu") instead of the character. Check every non-space character in
    `text` actually exists in the chosen font before using it."""
    try:
        font_obj = fitz.Font(fontfile=fontfile) if fontfile else fitz.Font(fontkey)
    except Exception:
        return False
    for ch in text:
        if ch.isspace():
            continue
        if not font_obj.has_glyph(ord(ch)):
            return False
    return True


def resolve_line_overlaps(insert_jobs, cover_rects):
    """When two different replacements land on the same visual line close
    together (e.g. a short bold label right next to a longer description
    on the same row), a replacement text that's LONGER than the text it
    replaced can run into the start of the next insertion and overlap it.

    This measures each insertion's actual rendered width (using the real
    font/size it will be drawn with, not a guess), and for any two
    insertions on the same line where the first would run past where the
    second starts, shifts the second (and anything after it on that line,
    cascading) to the right by exactly the overflow amount -- then widens
    that insertion's own cover rectangle to match its new extent, so the
    now-larger gap is still painted white rather than showing old text.
    """
    from collections import defaultdict

    lines = defaultdict(list)
    for job in insert_jobs:
        # Group by baseline, tolerating tiny rounding differences between
        # jobs that are genuinely on the same printed line.
        lines[round(job["y"])].append(job)

    gap = 1.0  # small breathing room between adjacent pieces of text
    for jobs in lines.values():
        if len(jobs) < 2:
            continue
        jobs.sort(key=lambda j: j["x"])
        for i in range(len(jobs) - 1):
            cur = jobs[i]
            font_obj = fitz.Font(fontfile=cur["fontfile"]) if cur["fontfile"] else fitz.Font(cur["fontkey"])
            cur_width = font_obj.text_length(cur["text"], fontsize=cur["size"])
            cur_end = cur["x"] + cur_width
            nxt = jobs[i + 1]
            if cur_end + gap <= nxt["x"]:
                continue
            shift = cur_end + gap - nxt["x"]
            nxt["x"] += shift
            nxt_font_obj = fitz.Font(fontfile=nxt["fontfile"]) if nxt["fontfile"] else fitz.Font(nxt["fontkey"])
            nxt_width = nxt_font_obj.text_length(nxt["text"], fontsize=nxt["size"])
            cover = cover_rects[nxt["cover_idx"]]
            cover.x1 = max(cover.x1, nxt["x"] + nxt_width + gap)



def find_near_miss_texts(spans, old):
    """A replace rule only matches text that is EXACTLY the same as `old`.
    If the PDF has a slightly different version of that text nearby (an
    extra word, a typo, different spacing -- e.g. the rule says "Aluclad
    Timber 68 PEFC Jointed Pine" but this page actually says "Aluclad
    Timber 68x80 PEFC Jointed Pine"), the rule silently skips it with no
    error at all, which is easy to miss.

    `old` is usually only PART of a span's text (e.g. it sits after a
    "System: " label), so we can't assume it starts at position 0 of the
    span. Instead: take a long leading chunk of `old` and check whether
    that chunk shows up anywhere inside a span's text that ISN'T an exact
    match -- a strong sign a near-duplicate slipped through. Returns the
    list of distinct near-miss texts found.
    """
    chunk_len = max(10, int(len(old) * 0.4))
    chunk = old[:chunk_len].lower()
    seen = set()
    variants = []
    for span in spans:
        text = span["text"].strip()
        if not text or old in text:
            continue
        if chunk in text.lower() and text not in seen:
            seen.add(text)
            variants.append(text)
    return variants


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
    not_found_warnings = []  # human-readable strings about rules that matched nothing

    # raw_hit_counts: every time a rule's "old" text was found on a page
    # this rule wasn't entirely gated out of (used to tell "genuinely not
    # found anywhere" apart from "found, but Only/Except correctly
    # restricted it away").
    raw_hit_counts = [0] * len(replace_rules)
    applied_hit_counts = [0] * len(replace_rules)
    raw_delete_hit_counts = [0] * len(delete_phrases)
    delete_resolved = [parse_delete_phrase(p["phrase"]) for p in delete_phrases]
    near_miss = {}  # rule_idx -> {variant_text: set(page_num, ...)}

    for pno in range(len(doc)):
        page = doc[pno]
        page_num = pno + 1
        spans = get_spans(page)
        pos_blocks = build_pos_blocks(spans, page.rect.y1)

        cover_rects = []
        insert_jobs = []  # list of dicts: x, y, text, fontkey, fontfile, size, cover_idx
        raw_hits = []  # collected across all rules before span-grouping (see below)

        for rule_idx, rule in enumerate(replace_rules):
            old, new = rule["old"], rule["new"]
            if not old:
                continue
            # Page-level gates -- these can be decided before even
            # searching the page, since they're page-number-based.
            if rule["except_pages"] and page_num in rule["except_pages"]:
                continue
            if rule["only_pages"] and page_num not in rule["only_pages"]:
                continue
            for rect in page.search_for(old):
                # PyMuPDF's search_for() matches case-INSENSITIVELY by
                # default, but these rules are documented (and promised in
                # the instructions template) to be case-sensitive -- e.g. a
                # rule for "AMSTERDAM F1" must not also catch "Amsterdam F1"
                # elsewhere on the same line. Verify the actual text in this
                # rect matches the rule's exact case before treating it as a
                # real hit; a case-differing match is silently skipped here
                # (not counted at all) so it doesn't inflate "not found"
                # bookkeeping either.
                if not _same_text_exact_case(page.get_textbox(rect), old):
                    continue
                raw_hit_counts[rule_idx] += 1

                # POS-number-based gates -- these need the actual match
                # location, since a single page can hold several POS items.
                if rule["only_pos"] or rule["except_pos"]:
                    pos_num = pos_number_for_rect(pos_blocks, rect)
                    if rule["only_pos"] and pos_num not in rule["only_pos"]:
                        continue
                    if rule["except_pos"] and pos_num in rule["except_pos"]:
                        continue

                applied_hit_counts[rule_idx] += 1
                raw_hits.append({"rect": rect, "old": old, "new": new, "size": rule["size"]})

            for variant_text in find_near_miss_texts(spans, old):
                near_miss.setdefault(rule_idx, {}).setdefault(variant_text, set()).add(page_num)

        # Two different rules can both match text that lives inside the
        # SAME underlying span (e.g. a summary line like "6. AMSTERDAM F1
        # - 5 quantity: - Hoppe Amsterdam F1 window handle, no key 303.75"
        # is one span, and separate rules target "AMSTERDAM F1" and "Hoppe
        # Amsterdam F1 window handle, no key" within it). Handling each
        # hit in isolation -- covering/rewriting from its match to the end
        # of the WHOLE span -- would make the earlier hit's rewrite
        # duplicate everything the later hit is separately (and correctly)
        # already handling, producing overlapping garbled text. Group hits
        # by their containing span and bound each one's rewritten region
        # to stop right before the next hit in that same span, instead of
        # running all the way to the span's end.
        pad = 0.4
        hits_by_span = {}
        standalone_hits = []
        for hit in raw_hits:
            span = find_containing_span(spans, hit["rect"])
            hit["span"] = span
            if span is None:
                standalone_hits.append(hit)
            else:
                hits_by_span.setdefault(id(span), []).append(hit)

        for hit in standalone_hits:
            print(f"  [warn] page {page_num}: could not find styling for "
                  f"'{hit['old']}' -- using a fallback font/position", file=sys.stderr)
            rect = hit["rect"]
            baseline_y = rect.y1 - (rect.y1 - rect.y0) * 0.2
            fontkey, fontfile = "helv", None
            size = hit["size"] or (rect.y1 - rect.y0) * 0.8
            combined_text = hit["new"]
            if not font_covers_text(fontkey, fontfile, combined_text):
                fontkey, fontfile = _base14_for_style(fontkey), None
            cover_rects.append(fitz.Rect(rect.x0 - pad, rect.y0 - pad, rect.x1 + pad, rect.y1 + pad))
            insert_jobs.append({
                "x": rect.x0, "y": baseline_y, "text": combined_text,
                "fontkey": fontkey, "fontfile": fontfile, "size": size,
                "cover_idx": len(cover_rects) - 1,
            })
            total_replaced += 1
            modified_pages.add(page_num)

        for span_hits in hits_by_span.values():
            span = span_hits[0]["span"]
            span_hits.sort(key=lambda h: h["rect"].x0)
            baseline_y = span["origin"][1]
            fontkey, fontfile = resolve_font(span["font"], font_files)
            span_x1 = span["bbox"][2]
            search_from = 0
            for i, hit in enumerate(span_hits):
                old, new, rect = hit["old"], hit["new"], hit["rect"]
                idx = span["text"].find(old, search_from)
                if idx == -1:
                    idx = span["text"].find(old)
                if i + 1 < len(span_hits):
                    next_idx = span["text"].find(span_hits[i + 1]["old"], idx + len(old))
                    suffix_end = next_idx if next_idx != -1 else len(span["text"])
                    cover_x1 = span_hits[i + 1]["rect"].x0
                else:
                    suffix_end = len(span["text"])
                    cover_x1 = span_x1
                suffix = span["text"][idx + len(old):suffix_end] if idx != -1 else ""
                search_from = suffix_end
                size = hit["size"] if hit["size"] else span["size"]
                combined_text = new + suffix

                # The document's embedded font is usually a SUBSET containing
                # only the glyphs the original file actually used. If the
                # replacement text needs a character that subset doesn't have
                # (a digit, an accent, a symbol...), using it anyway would
                # render as a blank box. Fall back to a full-coverage builtin
                # font (matching the same bold/italic style) in that case.
                use_fontkey, use_fontfile = fontkey, fontfile
                if not font_covers_text(use_fontkey, use_fontfile, combined_text):
                    print(f"  [warn] page {page_num}: font '{use_fontkey}' is missing a "
                          f"character needed for '{combined_text}' -- falling back to "
                          f"a built-in font", file=sys.stderr)
                    use_fontkey, use_fontfile = _base14_for_style(use_fontkey), None

                # Cover from the start of the old text through to the start
                # of the next hit in this span (or the end of the span if
                # this is the last hit) -- so any trailing text sharing the
                # span, e.g. "OTHER - 3 quantity:", gets reflowed after the
                # new text instead of being overlapped by it, without
                # stepping on territory another hit already owns.
                cover_rects.append(fitz.Rect(rect.x0 - pad, rect.y0 - pad, cover_x1 + pad, rect.y1 + pad))
                insert_jobs.append({
                    "x": rect.x0, "y": baseline_y, "text": combined_text,
                    "fontkey": use_fontkey, "fontfile": use_fontfile, "size": size,
                    "cover_idx": len(cover_rects) - 1,
                })
                total_replaced += 1
                modified_pages.add(page_num)

        for phrase_idx, phrase_rule in enumerate(delete_phrases):
            if phrase_rule["except_pages"] and page_num in phrase_rule["except_pages"]:
                continue
            if phrase_rule["only_pages"] and page_num not in phrase_rule["only_pages"]:
                continue
            search_text, whole_line = delete_resolved[phrase_idx]
            for rect in page.search_for(search_text):
                # Same case-sensitivity guard as the replace rules above.
                if not _same_text_exact_case(page.get_textbox(rect), search_text):
                    continue
                raw_delete_hit_counts[phrase_idx] += 1

                if phrase_rule["only_pos"] or phrase_rule["except_pos"]:
                    pos_num = pos_number_for_rect(pos_blocks, rect)
                    if phrase_rule["only_pos"] and pos_num not in phrase_rule["only_pos"]:
                        continue
                    if phrase_rule["except_pos"] and pos_num in phrase_rule["except_pos"]:
                        continue

                span = find_containing_span(spans, rect)
                pad = 0.4
                if whole_line and span is not None:
                    # The instruction said to remove the entire line this
                    # text lives on, not just the matched words -- e.g.
                    # "the whole line of 'U-value (W/m2K)'" means delete
                    # the whole "U-value (W/m2K)= 1.39" line.
                    lx0, ly0, lx1, ly1 = span.get("line_bbox", span["bbox"])
                    cover_rects.append(fitz.Rect(lx0 - pad, ly0 - pad, lx1 + pad, ly1 + pad))
                    total_deleted += 1
                    modified_pages.add(page_num)
                    continue
                if span is not None:
                    remainder = span["text"].replace(search_text, "", 1)
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

        resolve_line_overlaps(insert_jobs, cover_rects)

        # Non-destructive white overlay (rather than true redaction) --
        # redacting text can corrupt shared embedded font glyphs elsewhere
        # on the same page. The tradeoff: old text is visually hidden but
        # technically still present/extractable underneath.
        for r in cover_rects:
            page.draw_rect(r, color=None, fill=(1, 1, 1), fill_opacity=1, overlay=True)

        for job in insert_jobs:
            x, y, text, fontkey, fontfile, size = (
                job["x"], job["y"], job["text"], job["fontkey"], job["fontfile"], job["size"]
            )
            if fontfile:
                page.insert_text((x, y), text, fontsize=size, fontname=fontkey,
                                  fontfile=fontfile, color=(0, 0, 0))
            else:
                page.insert_text((x, y), text, fontsize=size, fontname=fontkey, color=(0, 0, 0))

        # Automated stand-in for a human eyeballing "the whole page, not
        # just the edited spot": re-check the page we just edited for any
        # inserted text unexpectedly overlapping something else.
        if insert_jobs:
            inserted_specs = [(job["x"], job["y"], job["text"]) for job in insert_jobs]
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

    # Flag any rule/phrase that never matched anywhere in the document --
    # almost always a typo, extra space, or the PDF wording being slightly
    # different from what was typed into the spreadsheet. Note this uses
    # raw_hit_counts (found at all, before Only/Except filtering) rather
    # than applied_hit_counts, so a rule correctly restricted down to zero
    # actual edits by its own Only/Except settings does NOT get flagged --
    # that's the sheet working as intended, not a mistake.
    for rule_idx, rule in enumerate(replace_rules):
        if raw_hit_counts[rule_idx] == 0:
            not_found_warnings.append(
                f"Replace rule '{rule['old']}' -> '{rule['new']}' was not "
                f"found anywhere in the PDF (outside any Only/Except pages) -- "
                f"double-check the spelling/spacing matches the PDF exactly."
            )
        variants = near_miss.get(rule_idx)
        if variants:
            for variant_text, pgs in variants.items():
                not_found_warnings.append(
                    f"Your rule '{rule['old']}' -> '{rule['new']}' left "
                    f"\"{variant_text}\" unchanged on page(s) {sorted(pgs)} "
                    f"-- it's close to your rule's text but not identical, "
                    f"so it didn't match. Add a separate row with this exact "
                    f"wording if it should change too."
                )
    for phrase_idx, phrase_rule in enumerate(delete_phrases):
        if raw_delete_hit_counts[phrase_idx] == 0:
            search_text, whole_line = delete_resolved[phrase_idx]
            not_found_warnings.append(
                f"Delete instruction '{phrase_rule['phrase']}' (looking for "
                f"\"{search_text}\") was not found anywhere in the PDF "
                f"(outside any Only/Except pages) -- double-check the "
                f"spelling/spacing matches the PDF exactly."
            )

    return total_replaced, total_deleted, sorted(modified_pages), overlap_warnings, not_found_warnings


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("pdf", help="Input PDF file")
    parser.add_argument("xlsx", help="Instructions .xlsx file")
    parser.add_argument("output", help="Output PDF path")
    parser.add_argument("--previews", default=None, help="Directory to save preview PNGs of modified pages")
    args = parser.parse_args()

    replaced, deleted, pages, overlap_warnings, not_found_warnings = process(
        args.pdf, args.xlsx, args.output, args.previews
    )
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
    if not_found_warnings:
        print("\n*** SOME RULES DID NOT MATCH ANYTHING -- check these: ***")
        for w in not_found_warnings:
            print(f"  - {w}")
