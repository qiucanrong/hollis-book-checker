"""Read inputs and write a marked .xlsx copy."""

import csv
import io
from collections import Counter
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, PatternFill, Side
from openpyxl.utils import get_column_letter

from hollis_api import SearchResult
from isbn_utils import normalize_isbn


FILLS = {
    "green": PatternFill("solid", fgColor="DDF1D8"),
    "red": PatternFill("solid", fgColor="F7D9D9"),
    "yellow": PatternFill("solid", fgColor="FFF0BD"),
}

THIN = Side(style="thin", color="B7B7B7")
ALL_BORDERS = Border(
    left=THIN,
    right=THIN,
    top=THIN,
    bottom=THIN,
)

OUTPUT_HEADER = "Title of Book Found"


def read_input(filename, contents):
    """Return a workbook; preserve CSV values as text."""
    extension = Path(filename).suffix.lower()

    if extension == ".xlsx":
        return load_workbook(io.BytesIO(contents))

    if extension == ".csv":
        try:
            text = contents.decode("utf-8-sig")
            rows = list(csv.reader(io.StringIO(text)))
        except (UnicodeError, csv.Error) as exc:
            raise ValueError(
                "CSV must be valid UTF-8."
            ) from exc

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Books"

        for row in rows:
            sheet.append(row)

        return workbook

    raise ValueError("Upload an .xlsx or UTF-8 .csv file.")


def column_options(sheet):
    """Identify columns by position, allowing duplicate or blank headers."""
    if sheet.max_row < 1:
        return {}

    options = {}

    for col in range(1, sheet.max_column + 1):
        header = sheet.cell(1, col).value

        if header is None and not any(
            sheet.cell(row, col).value is not None
            for row in range(2, sheet.max_row + 1)
        ):
            continue

        label = (
            str(header).strip()
            if header is not None
            else "(blank header)"
        )

        options[f"{get_column_letter(col)} — {label}"] = col

    return options


def data_rows(sheet, original_max_col):
    """Skip completely empty rows; treat row 1 as the header."""
    return [
        row
        for row in range(2, sheet.max_row + 1)
        if any(
            sheet.cell(row, col).value not in (None, "")
            for col in range(1, original_max_col + 1)
        )
    ]


def render_workbook(workbook, sheet_name, rows, results):
    sheet = workbook[sheet_name]

    # Insert the result column only after ISBN classification is complete.
    original_max_col = sheet.max_column
    sheet.insert_cols(1)

    output_col = 1
    last_col = original_max_col + 1

    header = sheet.cell(1, output_col)
    header.value = OUTPUT_HEADER

    # Original column A is now column B.
    if original_max_col:
        header._style = copy(sheet.cell(1, 2)._style)

    sheet.column_dimensions[get_column_letter(output_col)].width = 48

    counts = Counter()
    preview = []

    for row in rows:
        result = results[row]
        counts[result.status] += 1
        output = sheet.cell(row, output_col)

        if result.matches and result.status in ("red", "yellow"):
            if result.status == "red":
                match = result.matches[0]
                output.value = match.title

                if match.url:
                    output.hyperlink = match.url
                    output.font = copy(output.font)
                    output.font = output.font.copy(
                        color="0563C1",
                        underline="single",
                    )

            else:
                # Excel supports one hyperlink target per cell.
                # Include the multiple URLs as text.
                output.value = "\n".join(
                    f"{match.title} — {match.url}"
                    if match.url
                    else match.title
                    for match in result.matches
                )

                output.alignment = copy(output.alignment)
                output.alignment = output.alignment.copy(
                    wrap_text=True
                )

        # Apply colors and borders across the entire output row.
        for col in range(1, last_col + 1):
            cell = sheet.cell(row, col)

            if result.status in FILLS:
                cell.fill = copy(FILLS[result.status])

            cell.border = copy(ALL_BORDERS)

        preview.append({
            "Excel row": row,
            "Status": result.status,
            "Title of Book Found": output.value or "",
            "Review note": result.reason,
        })

    for col in range(1, last_col + 1):
        sheet.cell(1, col).border = copy(ALL_BORDERS)

    buffer = io.BytesIO()
    workbook.save(buffer)

    return buffer.getvalue(), counts, preview


def classify_rows(sheet, rows, isbn_col, client, progress=None):
    """Cache repeated ISBN searches within this processing run."""
    cache = {}
    results = {}

    for position, row in enumerate(rows, 1):
        isbn = normalize_isbn(sheet.cell(row, isbn_col).value)

        if isbn is None:
            result = SearchResult(
                "white",
                reason="Missing or invalid ISBN.",
            )

        elif isbn in cache:
            result = cache[isbn]

        else:
            try:
                result = client.search_isbn(isbn)

            except Exception as exc:
                # Avoid exposing raw request or credential diagnostics.
                from hollis_api import SearchError

                if isinstance(exc, SearchError):
                    result = SearchResult(
                        "white",
                        reason=str(exc),
                    )
                else:
                    result = SearchResult(
                        "white",
                        reason="Unexpected search error.",
                    )

            cache[isbn] = result

        results[row] = result

        if progress:
            progress(position, len(rows))

    return results