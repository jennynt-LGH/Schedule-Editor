"""
streamlit_app.py

A simple web front-end (upload button + download button) for
pdf_text_editor.py, meant to be deployed on Streamlit Community Cloud so
anyone with the link -- not just people comfortable with GitHub -- can use
the tool.

Nothing uploaded here gets saved into the GitHub repo: files are processed
in a temporary folder for the duration of the request and thrown away
afterwards.
"""

import os
import tempfile

import streamlit as st

from pdf_text_editor import process

st.set_page_config(page_title="Schedule PDF Editor", page_icon="\U0001F4C4")

st.title("Schedule PDF Editor")
st.write(
    "Upload a schedule PDF and an instructions spreadsheet. The tool finds "
    "and replaces (or deletes) the text you specify, matching the "
    "original font, size, and position as closely as possible."
)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template_instructions.xlsx")

if os.path.exists(TEMPLATE_PATH):
    with open(TEMPLATE_PATH, "rb") as f:
        template_bytes = f.read()
    st.download_button(
        "Download blank instructions template (.xlsx)",
        data=template_bytes,
        file_name="template_instructions.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("Fill this in with your replace/delete rules, save it, then upload it below.")

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("PDF file", type=["pdf"])
with col2:
    xlsx_file = st.file_uploader("Instructions spreadsheet (.xlsx)", type=["xlsx"])

with st.expander("What should the instructions spreadsheet look like?"):
    st.write(
        "First sheet, two sections:\n\n"
        "**Replace table** -- a row containing the header `Replace from "
        "this` starts it, with columns for `to this`, `Exception` "
        "(e.g. `page 9`, or `-`/blank for none), and `Font size of new "
        "text`.\n\n"
        "**Delete table** -- a row containing `Delete these words from "
        "PDF` starts it, followed by one phrase per row to remove.\n\n"
        "A blank row ends each table."
    )

process_clicked = st.button(
    "Process", type="primary", disabled=not (pdf_file and xlsx_file)
)

if not (pdf_file and xlsx_file):
    st.info("Upload both files above to enable the Process button.")

if process_clicked:
    with tempfile.TemporaryDirectory() as workdir:
        pdf_path = os.path.join(workdir, "input.pdf")
        xlsx_path = os.path.join(workdir, "instructions.xlsx")
        output_path = os.path.join(workdir, "output.pdf")
        preview_dir = os.path.join(workdir, "previews")

        with open(pdf_path, "wb") as f:
            f.write(pdf_file.getbuffer())
        with open(xlsx_path, "wb") as f:
            f.write(xlsx_file.getbuffer())

        try:
            with st.spinner("Processing..."):
                replaced, deleted, pages, overlap_warnings = process(
                    pdf_path, xlsx_path, output_path, preview_dir
                )
        except Exception as e:
            st.error(f"Something went wrong: {e}")
        else:
            st.success(
                f"Done -- {replaced} replacement(s), {deleted} deletion(s), "
                f"across page(s) {pages if pages else 'none'}."
            )

            with open(output_path, "rb") as f:
                pdf_bytes = f.read()

            out_name = os.path.splitext(pdf_file.name)[0] + "_updated.pdf"
            st.download_button(
                "Download finished PDF",
                data=pdf_bytes,
                file_name=out_name,
                mime="application/pdf",
            )

            if overlap_warnings:
                st.warning(
                    "The tool ran its own check of the whole page (not just "
                    "the edited spot) and found something that may need a "
                    "closer look before you use this file:"
                )
                for pno, page_warnings in overlap_warnings.items():
                    for w in page_warnings:
                        st.write(f"- Page {pno}: {w}")

            if pages:
                st.subheader("Preview of changed pages")
                st.caption(
                    "Quick visual check to catch any unintended overlap "
                    "before you send the file anywhere."
                )
                for pno in pages:
                    preview_path = os.path.join(preview_dir, f"page{pno}_preview.png")
                    if os.path.exists(preview_path):
                        st.image(preview_path, caption=f"Page {pno}")
