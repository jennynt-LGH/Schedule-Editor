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

    "Replace from this" can list several exact alternatives for the SAME
    "to this" text, separated by "|" -- e.g. "jonied onsite | joined
    onsiet" both becoming "joined onsite". This is for different known
    ways the same intended text can turn up broken (matching is always
    an exact literal substring, never a pattern/wildcard, so each
    alternative has to be spelled out), not for genuinely different
    replacements -- those still need their own separate row.

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

                # "Replace from this" can list several exact alternatives,
                # separated by "|", that should all become the SAME "to
                # this" text -- e.g. two different known-broken spellings
                # of the same word (a PDF-generator glyph-ordering defect
                # can scramble a word differently in different contexts)
                # that both need to end up as the same correct text.
                # Since matching is always an exact literal substring (no
                # wildcards), each alternative becomes its own internal
                # rule here, but they're written as one row for the user.
                old_alternatives = [alt.strip() for alt in str(old).split("|")]
                old_alternatives = [alt for alt in old_alternatives if alt]
                row_group_id = i  # shared by every alternative from this one row

                for old_alt in old_alternatives:
                    replace_rules.append({
                        "old": old_alt,
                        "new": "" if new is None else str(new).strip(),
                        "only_pages": only_pages,
                        "only_pos": only_pos,
                        "except_pages": except_pages,
                        "except_pos": except_pos,
                        "size": size,
                        "row_group_id": row_group_id,
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
    """Extract every text span on the page, using PyMuPDF's "rawdict" mode
    so each span carries its individual characters' bounding boxes (not
    just the span's own overall bbox). This lets us locate a substring's
    exact rectangle directly from the page's own extracted text, instead
    of re-querying the page with page.search_for() -- which, empirically,
    can occasionally fail to find text that unambiguously exists (this
    surfaced on one particular summary line where search_for() silently
    returned no match at all for a phrase clearly present in the page's
    own text). Searching within the extracted text we already trust
    removes that failure mode entirely.

    Each span also gets a "line_idx" -- a per-page counter identifying
    which of PyMuPDF's own "line" groupings it belongs to -- so spans
    that make up one continuous printed line can be found and searched
    together (see build_lines / search_text_in_lines), even when that
    line happens to be split into several spans. It also gets a
    "block_idx", identifying which of PyMuPDF's own "block" groupings
    (a level above "line") it belongs to -- some PDF generators draw a
    whole multi-line attribute paragraph as one shared block/text object,
    and this lets code elsewhere detect "these two edits fall in the same
    underlying block" to avoid a redaction-corruption failure mode that's
    specific to editing such a block more than once.
    """
    spans = []
    d = page.get_text("rawdict")
    line_idx = 0
    for block_idx, block in enumerate(d["blocks"]):
        if block.get("type") != 0:  # skip image blocks
            continue
        for line in block.get("lines", []):
            line_bbox = line["bbox"]
            for span in line["spans"]:
                span = dict(span)
                span["line_bbox"] = line_bbox
                span["line_idx"] = line_idx
                span["block_idx"] = block_idx
                chars = span.get("chars", [])
                span["text"] = "".join(c["c"] for c in chars)
                span["_chars"] = chars
                spans.append(span)
            line_idx += 1
    return spans


def build_lines(spans):
    """Group spans back into the printed lines PyMuPDF originally split
    them from (via the line_idx get_spans tagged them with), and build
    one continuous text string per line by concatenating that line's
    spans left-to-right -- along with a per-character map back to
    (span, local_char_index).

    This exists because a single printed line isn't always one span: a
    PDF generator will sometimes need a different embedded font for just
    one character in the middle of an otherwise plain line -- most often
    a diacritic (e.g. "osłonek") missing from the main subset font --
    which silently splits that line into 2 or 3 spans even though it
    looks completely uniform. Searching only within individual spans (as
    a plain page-text search effectively does) will never find a phrase
    that happens to cross such a split; searching each line's full
    reassembled text does.
    """
    by_line = {}
    for span in spans:
        by_line.setdefault(span["line_idx"], []).append(span)

    lines = []
    for line_spans in by_line.values():
        line_spans.sort(key=lambda s: s["bbox"][0])
        text_parts = []
        char_owners = []  # parallel to the concatenated text below
        for span in line_spans:
            for local_idx in range(len(span["_chars"])):
                text_parts.append(span["_chars"][local_idx]["c"])
                char_owners.append((span, local_idx))
        lines.append({"text": "".join(text_parts), "char_owners": char_owners})
    return lines


def search_text_in_lines(lines, text):
    """Find every exact, case-sensitive occurrence of `text` within each
    line's full reassembled text (see build_lines) -- so a match can span
    more than one underlying PDF span/font-run. Returns a list of dicts:
    rect (built from the matched characters' own bounding boxes, across
    however many spans they came from), the "primary" span the match
    STARTS in (used for the replacement's font/size/baseline -- the
    common case is a match entirely inside one span anyway, where this
    is simply that span), the line, and the match's [start, end) character
    offsets within that line's text (used later to bound how far a
    suffix can be reflowed without overrunning a different, later hit on
    the same line)."""
    results = []
    for line in lines:
        line_text = line["text"]
        owners = line["char_owners"]
        start = 0
        while True:
            idx = line_text.find(text, start)
            if idx == -1:
                break
            end = idx + len(text)
            matched_owners = owners[idx:end]
            if matched_owners:
                boxes = [owners[i][0]["_chars"][owners[i][1]]["bbox"] for i in range(idx, end)]
                x0 = min(b[0] for b in boxes)
                y0 = min(b[1] for b in boxes)
                x1 = max(b[2] for b in boxes)
                y1 = max(b[3] for b in boxes)
                results.append({
                    "rect": fitz.Rect(x0, y0, x1, y1),
                    "span": matched_owners[0][0],
                    "line": line,
                    "start": idx,
                    "end": end,
                })
            start = end
    return results


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
    font name (subset prefix like 'ABCDEF+' stripped) -- as a LIST of
    every distinct subset found under that name, not just one.

    PDF generators very commonly re-subset the same font independently
    per page (or per text block), each subset embedding only the glyphs
    THAT one instance actually uses. Two pages' "Arial", for example, can
    be two genuinely different files with different glyph coverage --
    one might include a Polish diacritic the other doesn't. Keeping only
    the last one seen (as an earlier version of this function did) meant
    a later page's lookup of "Arial" could silently get a DIFFERENT,
    incomplete subset that's missing a character the current text
    actually needs, even though some other subset earlier in the
    document has it. Keeping every subset and letting the caller check
    actual glyph coverage (see resolve_font_for_text) avoids that.
    """
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
            path = os.path.join(workdir, f"{base_name}.{xref}.{ext or 'ttf'}")
            with open(path, "wb") as fh:
                fh.write(buf)
            font_files.setdefault(base_name, []).append(path)
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


def resolve_font_for_text(span_font_name, font_files, text):
    """Return (fontname, fontfile_or_None) for insert_text(), specifically
    chosen so the returned font actually covers every character in `text`
    -- not just assumed to, based on name alone.

    font_files maps a base font name to a LIST of every distinct embedded
    subset found under that name anywhere in the document (see
    extract_all_fonts for why there can be more than one). Different
    subsets of the "same" font can have different glyph coverage, so this
    tries each one against the actual text being inserted here and uses
    the first that works, rather than an arbitrary one that happens to be
    missing a character this particular insertion needs. Falls back to a
    built-in base-14 font (matching bold/italic style) only if none of
    the embedded subsets found under this name cover the text either.
    """
    name = span_font_name or ""
    if "+" in name:
        name = name.split("+", 1)[1]
    for candidate_path in font_files.get(name, []):
        if font_covers_text(name, candidate_path, text):
            return name, candidate_path
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

    This is a backstop for cases the proactive line-flow positioning in
    process() doesn't cover (e.g. a fallback/standalone insertion whose
    baseline happens to coincide with another line). It only nudges
    something when there's genuine overlap -- not a plain "settle any two
    insertions on any line a bit apart", since the calling code has
    already positioned suffix-linked hits exactly flush with no gap on
    purpose, and adding one back in here would silently reintroduce the
    "unwanted extra space" this whole approach exists to avoid.
    """
    from collections import defaultdict

    lines = defaultdict(list)
    for job in insert_jobs:
        # Group by baseline, tolerating tiny rounding differences between
        # jobs that are genuinely on the same printed line.
        lines[round(job["y"])].append(job)

    epsilon = 0.05  # floating-point safety margin only, not a visible gap
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
            if cur_end <= nxt["x"] + epsilon:
                continue
            shift = cur_end - nxt["x"]
            nxt["x"] += shift
            nxt_font_obj = fitz.Font(fontfile=nxt["fontfile"]) if nxt["fontfile"] else fitz.Font(nxt["fontkey"])
            nxt_width = nxt_font_obj.text_length(nxt["text"], fontsize=nxt["size"])
            cover = cover_rects[nxt["cover_idx"]]
            cover.x1 = max(cover.x1, nxt["x"] + nxt_width)



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

    num_pages = len(doc)
    for pno in range(num_pages):
        page = doc[pno]
        page_num = pno + 1
        spans = get_spans(page)
        lines = build_lines(spans)
        pos_blocks = build_pos_blocks(spans, page.rect.y1)

        redact_rects = []  # tight rects around the EXACT old text being
                            # removed/replaced -- these get TRUE redaction
        mask_rects = []    # rects covering only REPOSITIONING territory
                            # (the gap a shorter/longer replacement opens or
                            # closes, or unmatched text being reflowed) --
                            # these get the older, non-destructive white
                            # overlay instead. Splitting the two apart keeps
                            # redaction's footprint as small as possible:
                            # some PDF generators draw a whole multi-line
                            # attribute block as one internal text object,
                            # and redacting a large rect that merely touches
                            # such a block can wipe out far more than
                            # intended. A tight redaction rect around just
                            # the actual old word/phrase is much less likely
                            # to trigger that.
        insert_jobs = []  # list of dicts: x, y, text, fontkey, fontfile, size, cover_idx (-> mask_rects)
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
            for hit in search_text_in_lines(lines, old):
                rect = hit["rect"]
                # search_text_in_lines() only ever returns exact,
                # case-sensitive matches (it searches the literal
                # extracted text directly) -- so no separate case check
                # is needed here the way page.search_for() used to need.
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
                raw_hits.append({
                    "rect": rect, "old": old, "new": new, "size": rule["size"],
                    "span": hit["span"], "line": hit["line"],
                    "start": hit["start"], "end": hit["end"],
                })

            for variant_text in find_near_miss_texts(spans, old):
                near_miss.setdefault(rule_idx, {}).setdefault(variant_text, set()).add(page_num)

        # Two different rules can both match text that lives on the SAME
        # underlying printed line (e.g. a summary line like "6. AMSTERDAM
        # F1 - 5 quantity: - Hoppe Amsterdam F1 window handle, no key
        # 303.75", where separate rules target "AMSTERDAM F1" and "Hoppe
        # Amsterdam F1 window handle, no key" within it). Handling each
        # hit in isolation -- covering/rewriting from its match to the end
        # of the WHOLE line -- would make the earlier hit's rewrite
        # duplicate everything the later hit is separately (and correctly)
        # already handling, producing overlapping garbled text. Group hits
        # by their line and bound each one's rewritten region to stop
        # right before the next hit on that same line, instead of running
        # all the way to the line's end.
        pad = 0.4
        hits_by_line = {}
        standalone_hits = []
        for hit in raw_hits:
            span = hit["span"]
            if span is None:
                standalone_hits.append(hit)
            else:
                hits_by_line.setdefault(id(hit["line"]), []).append(hit)

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
            redact_rects.append(fitz.Rect(rect.x0 - pad, rect.y0 - pad, rect.x1 + pad, rect.y1 + pad))
            # No containing span was found for this hit (that's what makes
            # it "standalone"), so there's no block_idx to key off of --
            # fall back to the POS-block heuristic instead.
            redact_idx = len(redact_rects) - 1
            # No reflow gap here (no suffix to bridge), but insert_jobs
            # always tracks a mask_rects entry for resolve_line_overlaps to
            # extend if it ever needs to -- start it zero-width.
            mask_rects.append(fitz.Rect(rect.x1 - pad, rect.y0 - pad, rect.x1 + pad, rect.y1 + pad))
            insert_jobs.append({
                "x": rect.x0, "y": baseline_y, "text": combined_text,
                "fontkey": fontkey, "fontfile": fontfile, "size": size,
                "cover_idx": len(mask_rects) - 1,
                "redact_idx": redact_idx,
                "mask_indices": [len(mask_rects) - 1],
            })
            total_replaced += 1
            modified_pages.add(page_num)

        for line_hits in hits_by_line.values():
            line = line_hits[0]["line"]
            line_hits.sort(key=lambda h: h["rect"].x0)
            flow_x = None  # where the previous hit's rendered text actually ended
            for i, hit in enumerate(line_hits):
                span = hit["span"]
                rect = hit["rect"]
                baseline_y = span["origin"][1]
                idx, end = hit["start"], hit["end"]
                if i + 1 < len(line_hits):
                    suffix_end = line_hits[i + 1]["start"]
                    cover_x1 = line_hits[i + 1]["rect"].x0
                else:
                    suffix_end = len(line["text"])
                    cover_x1 = span["line_bbox"][2]
                suffix = line["text"][end:suffix_end]
                size = hit["size"] if hit["size"] else span["size"]
                combined_text = hit["new"] + suffix

                # The document's embedded font is usually a SUBSET containing
                # only the glyphs the original file actually used, and the
                # SAME font name can have several different subsets across
                # the document with different glyph coverage (see
                # extract_all_fonts). Pick whichever actual subset -- or,
                # failing that, a full-coverage builtin font matching the
                # same bold/italic style -- genuinely covers everything in
                # this specific combined_text (the replacement plus
                # whatever original suffix text is being reflowed with it).
                use_fontkey, use_fontfile = resolve_font_for_text(span["font"], font_files, combined_text)
                if use_fontfile is None and not font_covers_text(use_fontkey, use_fontfile, combined_text):
                    print(f"  [warn] page {page_num}: no available font (embedded or "
                          f"built-in) covers every character needed for '{combined_text}' "
                          f"-- some characters may not render correctly", file=sys.stderr)

                # Where to actually draw this hit's text. The FIRST hit on a
                # line keeps its original position. Every hit after that
                # starts exactly where the PREVIOUS hit's own inserted text
                # (combined_text, suffix included) actually ends -- not at
                # its own old position -- because that suffix already
                # contains everything that used to sit between the two
                # hits (e.g. ": "). If we left it at its old position
                # instead, a replacement that's SHORTER than what it
                # replaced (e.g. "Kolor osłonek" -> "Hinge caps") would
                # leave a stretch of dead white space where the extra
                # length used to be, and a LONGER replacement would
                # overlap the next hit -- this single rule fixes both.
                x = rect.x0 if flow_x is None else flow_x
                font_obj = fitz.Font(fontfile=use_fontfile) if use_fontfile else fitz.Font(use_fontkey)
                rendered_width = font_obj.text_length(combined_text, fontsize=size)
                flow_x = x + rendered_width

                # Cover from the start of the old text (or, if this hit's
                # new position was pulled left of that, from the new
                # position instead) through to the start of the next hit on
                # this line, or however far the new text actually reaches
                # if that's further right than the next old hit started --
                # so any trailing text sharing the line, e.g. "OTHER - 3
                # quantity:", gets reflowed after the new text instead of
                # being overlapped by it, without stepping on territory
                # another hit already owns.
                cover_x0 = min(rect.x0, x)
                cover_x1 = max(cover_x1, flow_x)

                # The OLD TEXT ITSELF gets a tight, exact-fit true-redaction
                # rect (see the note on redact_rects/mask_rects above).
                # Everything else in the covered span -- the gap before or
                # after it that only exists to reflow surrounding text --
                # is pure repositioning, not sensitive replaced content, so
                # it gets the safer non-destructive mask instead.
                redact_rects.append(fitz.Rect(rect.x0 - pad, rect.y0 - pad, rect.x1 + pad, rect.y1 + pad))
                redact_idx = len(redact_rects) - 1
                mask_rects.append(fitz.Rect(rect.x1 - pad, rect.y0 - pad, cover_x1 + pad, rect.y1 + pad))
                mask_idx = len(mask_rects) - 1
                mask_indices = [mask_idx]
                if cover_x0 < rect.x0 - 1e-6:
                    mask_rects.append(fitz.Rect(cover_x0 - pad, rect.y0 - pad, rect.x0 + pad, rect.y1 + pad))
                    mask_indices.append(len(mask_rects) - 1)

                insert_jobs.append({
                    "x": x, "y": baseline_y, "text": combined_text,
                    "fontkey": use_fontkey, "fontfile": use_fontfile, "size": size,
                    "cover_idx": mask_idx,
                    "redact_idx": redact_idx,
                    "mask_indices": mask_indices,
                })
                total_replaced += 1
                modified_pages.add(page_num)

        for phrase_idx, phrase_rule in enumerate(delete_phrases):
            if phrase_rule["except_pages"] and page_num in phrase_rule["except_pages"]:
                continue
            if phrase_rule["only_pages"] and page_num not in phrase_rule["only_pages"]:
                continue
            search_text, whole_line = delete_resolved[phrase_idx]
            for hit in search_text_in_lines(lines, search_text):
                rect, span = hit["rect"], hit["span"]
                # search_text_in_lines() only returns exact, case-sensitive
                # matches, so no separate case check is needed here.
                raw_delete_hit_counts[phrase_idx] += 1

                if phrase_rule["only_pos"] or phrase_rule["except_pos"]:
                    pos_num = pos_number_for_rect(pos_blocks, rect)
                    if phrase_rule["only_pos"] and pos_num not in phrase_rule["only_pos"]:
                        continue
                    if phrase_rule["except_pos"] and pos_num in phrase_rule["except_pos"]:
                        continue

                pad = 0.4
                # Delete phrases always genuinely remove content (there's no
                # replacement text to reflow around), so everything here
                # goes straight to true redaction.
                if whole_line and span is not None:
                    # The instruction said to remove the entire line this
                    # text lives on, not just the matched words -- e.g.
                    # "the whole line of 'U-value (W/m2K)'" means delete
                    # the whole "U-value (W/m2K)= 1.39" line.
                    lx0, ly0, lx1, ly1 = span.get("line_bbox", span["bbox"])
                    redact_rects.append(fitz.Rect(lx0 - pad, ly0 - pad, lx1 + pad, ly1 + pad))
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
                        redact_rects.append(fitz.Rect(bx0 - pad, by0 - pad, bx1 + pad, by1 + pad))
                        total_deleted += 1
                        modified_pages.add(page_num)
                        continue
                redact_rects.append(fitz.Rect(rect.x0 - pad, rect.y0 - pad, rect.x1 + pad, rect.y1 + pad))
                total_deleted += 1
                modified_pages.add(page_num)

        if not redact_rects and not mask_rects:
            continue

        resolve_line_overlaps(insert_jobs, mask_rects)

        # Map each redact_rects index to the insert_job (if any) it belongs
        # to, so each hit's full lifecycle -- redact its old text, mask its
        # reflow gap, draw its new text -- can be finished completely
        # before moving to the next hit.
        job_by_redact_idx = {}
        for job in insert_jobs:
            ridx = job.get("redact_idx")
            if ridx is not None:
                job_by_redact_idx[ridx] = job

        def _draw_insert(page, job):
            x, y, text, fontkey, fontfile, size = (
                job["x"], job["y"], job["text"], job["fontkey"], job["fontfile"], job["size"]
            )
            if fontfile:
                page.insert_text((x, y), text, fontsize=size, fontname=fontkey,
                                  fontfile=fontfile, color=(0, 0, 0))
            else:
                page.insert_text((x, y), text, fontsize=size, fontname=fontkey, color=(0, 0, 0))

        # The old text actually being replaced/deleted gets TRUE redaction
        # -- removed from the page's content stream, not just painted over
        # -- so a PDF viewer's search/copy no longer finds it. Kept as
        # tight, minimal rects (see redact_rects/mask_rects note above) to
        # limit how much surrounding content a redaction could ever touch.
        # images=PDF_REDACT_IMAGE_NONE keeps this scoped to text only, so
        # it can't affect the window/door diagrams even if a rect happens
        # to sit close to one.
        #
        # Each hit's redact -> mask -> insert is applied as one complete
        # unit, in sequence, before starting the next hit's -- rather than
        # batching every redaction on the page into one add-all-then-
        # apply-once call. On some documents, particular attribute blocks
        # that get edited in more than one place have been seen to develop
        # corruption in unrelated, unedited nearby text after true
        # redaction -- if that happens on a document you're working with,
        # it's worth flagging so this can be revisited.
        for i, r in enumerate(redact_rects):
            page.add_redact_annot(r, fill=(1, 1, 1))
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
            job = job_by_redact_idx.get(i)
            if job is not None:
                for midx in job["mask_indices"]:
                    page.draw_rect(mask_rects[midx], color=None, fill=(1, 1, 1),
                                    fill_opacity=1, overlay=True)
                _draw_insert(page, job)

        # Any mask rects not tied to a specific redaction (there currently
        # are none by construction, but this keeps behavior correct if that
        # ever changes) still get painted.
        drawn_mask_indices = {midx for job in insert_jobs for midx in job["mask_indices"]}
        for midx, r in enumerate(mask_rects):
            if midx not in drawn_mask_indices:
                page.draw_rect(r, color=None, fill=(1, 1, 1), fill_opacity=1, overlay=True)

        # Any insert job that somehow isn't tied to a redaction (shouldn't
        # happen given how insert_jobs are built, but keeps behavior
        # correct if that ever changes) still gets drawn.
        for job in insert_jobs:
            if job.get("redact_idx") is None:
                _draw_insert(page, job)

        # Automated stand-in for a human eyeballing "the whole page, not
        # just the edited spot": re-check the page we just edited for any
        # inserted text unexpectedly overlapping something else.
        if insert_jobs:
            inserted_specs = [(job["x"], job["y"], job["text"]) for job in insert_jobs]
            page_warnings = find_overlap_warnings(page, inserted_specs, redact_rects + mask_rects)
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
    #
    # A row that listed several "|" alternatives is only flagged if NONE
    # of them matched anything -- finding some alternatives but not
    # others in a given document is exactly what listing alternatives is
    # FOR (different documents can have different known-broken spellings
    # of the same word), not a mistake to warn about.
    warned_groups = set()
    for rule_idx, rule in enumerate(replace_rules):
        group_id = rule.get("row_group_id")
        if raw_hit_counts[rule_idx] == 0:
            sibling_idxs = [
                j for j, r in enumerate(replace_rules)
                if r.get("row_group_id") == group_id
            ]
            if all(raw_hit_counts[j] == 0 for j in sibling_idxs) and group_id not in warned_groups:
                warned_groups.add(group_id)
                if len(sibling_idxs) > 1:
                    alternatives = " | ".join(replace_rules[j]["old"] for j in sibling_idxs)
                    not_found_warnings.append(
                        f"None of the alternatives '{alternatives}' -> '{rule['new']}' "
                        f"were found anywhere in the PDF (outside any Only/Except pages) -- "
                        f"double-check the spelling/spacing matches the PDF exactly."
                    )
                else:
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
