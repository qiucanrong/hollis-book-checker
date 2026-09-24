"""Streamlit interface for HOLLIS ISBN purchase screening."""

import hashlib
import logging
import os
from pathlib import Path

import streamlit as st

from hollis_api import HollisClient
from spreadsheet_processor import (
    classify_rows,
    column_options,
    data_rows,
    read_input,
    render_workbook,
)


logging.basicConfig(
    level=getattr(
        logging,
        os.getenv("HOLLIS_LOG_LEVEL", "WARNING").upper(),
        logging.WARNING,
    ),
    format="%(levelname)s %(name)s: %(message)s",
)

st.set_page_config(
    page_title="HOLLIS ISBN Checker",
    layout="wide",
)

st.title("HOLLIS ISBN Purchase Screening")

st.write(
    "Upload a book list, select its ISBN and title columns, "
    "and download an Excel copy with rows marked for review."
)

st.caption(
    "Green: no verified match · Red: one match · "
    "Yellow: multiple matches · Uncolored: manual review"
)

uploaded = st.file_uploader(
    "Book spreadsheet",
    type=["xlsx", "csv"],
)

if uploaded is None:
    st.stop()

contents = uploaded.getvalue()

try:
    workbook = read_input(uploaded.name, contents)
except Exception:
    st.error(
        "Could not read this spreadsheet. "
        "Use an .xlsx or UTF-8 .csv file."
    )
    st.stop()

sheet_name = st.selectbox(
    "Worksheet",
    workbook.sheetnames,
)

sheet = workbook[sheet_name]
options = column_options(sheet)

if len(options) < 2:
    st.error(
        "The first row needs identifiable ISBN and title columns."
    )
    st.stop()

labels = list(options)

isbn_default = next(
    (
        i for i, name in enumerate(labels)
        if "isbn" in name.lower()
    ),
    0,
)

title_default = next(
    (
        i for i, name in enumerate(labels)
        if "title" in name.lower()
    ),
    0,
)

isbn_label = st.selectbox(
    "ISBN column",
    labels,
    index=isbn_default,
)

title_label = st.selectbox(
    "Original title column",
    labels,
    index=title_default,
)

environment = st.selectbox(
    "HOLLIS environment",
    ["Sandbox", "Production"],
)

st.caption(
    "Use the API key issued for the selected environment."
)

isbn_col = options[isbn_label]
title_col = options[title_label]

if isbn_col == title_col:
    st.warning("Select different ISBN and title columns.")
    st.stop()

rows = data_rows(sheet, sheet.max_column)
st.write(f"{len(rows)} nonempty book rows found.")

signature = (
    hashlib.sha256(contents).hexdigest(),
    uploaded.name,
    sheet_name,
    isbn_col,
    title_col,
    environment,
)

if st.session_state.get("result_signature") != signature:
    st.session_state.pop("completed_result", None)

if st.button("Search HOLLIS", disabled=not rows):
    try:
        api_key = st.secrets["HOLLIS_API_KEY"]

        if not isinstance(api_key, str) or not api_key.strip():
            raise KeyError("HOLLIS_API_KEY")

    except (KeyError, FileNotFoundError):
        st.error(
            "HOLLIS_API_KEY is missing from Streamlit secrets."
        )
        st.stop()

    bar = st.progress(0)
    message = st.empty()

    def update_progress(done, total):
        bar.progress(done / total)
        message.write(f"{done} of {total} rows processed.")

    client = HollisClient(
        api_key.strip(),
        environment,
    )

    try:
        results = classify_rows(
            sheet,
            rows,
            isbn_col,
            client,
            update_progress,
        )
    finally:
        client.close()

    # Insert the output column after all ISBN values have been read.
    output, counts, preview = render_workbook(
        workbook,
        sheet_name,
        rows,
        results,
    )

    filename = f"Searched_{Path(uploaded.name).stem}.xlsx"

    st.session_state["result_signature"] = signature
    st.session_state["completed_result"] = (
        output,
        filename,
        dict(counts),
        preview,
    )

completed = st.session_state.get("completed_result")

if (
    completed
    and st.session_state.get("result_signature") == signature
):
    output, filename, counts, preview = completed

    st.subheader("Results")
    columns = st.columns(5)

    labels_and_counts = [
        ("Rows", len(rows)),
        ("Purchase candidates", counts.get("green", 0)),
        ("Already in HOLLIS", counts.get("red", 0)),
        ("Multiple records", counts.get("yellow", 0)),
        ("Manual review", counts.get("white", 0)),
    ]

    for column, (label, value) in zip(
        columns, labels_and_counts
    ):
        column.metric(label, value)

    st.dataframe(
        preview,
        hide_index=True,
        use_container_width=True,
    )

    st.download_button(
        "Download checked workbook",
        data=output,
        file_name=filename,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )