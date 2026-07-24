# PDF text editor (runs automatically on GitHub)

This little tool does what Claude did for you manually: it takes a PDF and a
list of find-and-replace / delete instructions (in an Excel file), and
produces an updated PDF that matches the original font, size, and position
as closely as possible.

Instead of asking Claude each time, GitHub itself will run the code for you
whenever you upload files. No installing anything, no hosting, nothing to
maintain.

## One-time setup

1. Create a new GitHub repository (or use an empty one).
2. Upload every file and folder from this bundle into it, **keeping the
   folder structure exactly as it is**:
   - `pdf_text_editor.py`
   - `requirements.txt`
   - `.github/workflows/process-pdf.yml`
   - `input/` (empty folder, this is where you'll drop files)
   - `output/` (empty folder, this is where results appear)
   - `example_instructions.xlsx` (just a reference/example — see below)

   On github.com you can do this with the "Add file" → "Upload files"
   button, dragging in the whole folder.

That's it. Nothing else to install or configure.

## Every time you want to process a PDF

1. Go to the `input/` folder in your repo.
2. Upload your PDF and your instructions `.xlsx` file there (again via
   "Add file" → "Upload files"), and commit.
3. Click the **Actions** tab at the top of the repo. You'll see a run start
   automatically — it usually finishes in under a minute. A green
   checkmark means it worked; a red X means something went wrong (click
   into it to read the error, usually a missing PDF or xlsx file).
4. Once it's green, go back to the `output/` folder. You'll find:
   - `<yourfile>_updated.pdf` — the finished PDF
   - `<yourfile>_previews/` — a picture of every page that was changed,
     so you can quickly eyeball that nothing overlaps oddly before you
     use the real file
5. Download what you need. You can then delete the files from `input/`
   and `output/` any time to keep the repo tidy — this has no real
   storage cost either way (see note below).

## The instructions Excel file format

Open `example_instructions.xlsx` in this bundle to see the expected layout.
It has two sections on the first sheet:

**Replace table** — a row with the header `Replace from this` starts it:

| Replace from this | to this | Exception | Font size of new text |
|---|---|---|---|
| old phrase | new phrase | page 9 | 12.08 |

- `Exception` can be blank, `-`, or something like `page 9` — any page
  numbers mentioned there are skipped for that rule.
- `Font size of new text` is the point size to use for the replacement.
- The new text automatically matches the original's font and style
  (regular / bold / italic) and lines up in the same spot — you don't
  need to specify that separately.

**Delete table** — a row containing `Delete these words from PDF` starts
it, followed by one phrase per row to remove entirely.

Leave a blank row to end each table (see the example file).

**Notes and guide text are safe to leave in the sheet.** The tool only
starts reading once it finds the exact header row (`Replace from this`),
and each table stops at its first blank row — so anything above the
header (a title, a "how to fill this in" guide row) or below that
closing blank row (a reminder note, etc.) is simply ignored, never read
as data. That's why the template has a guide row and a closing reminder
in it and still works.

If you're using the web app, there's a "Download blank instructions
template" button that gives you this same file to fill in.

## Good to know / limitations

- **This covers text, it doesn't erase it.** The tool draws a white box
  over the old text and writes the new text on top, rather than truly
  deleting the old text from the file. Visually it's correct, but if
  someone copies text out of the PDF or searches it, traces of the old
  text can still turn up underneath. Genuinely removing it (true
  redaction) was tried first but occasionally corrupted unrelated text
  elsewhere on the page, so this safer approach is the default.
- **Always check the preview images** in `output/<file>_previews/` before
  sending the final PDF anywhere — this is the same visual check Claude
  does by hand, just automated here.
- The tool expects the text you're searching for to exist as real,
  selectable text in the PDF (not text inside a scanned image).
- If a phrase appears in a spot the tool can't confidently match to a
  font (rare), it'll fall back to a plain font at a best-guess size and
  print a warning in the Actions log — worth a closer look at the preview
  in that case.
