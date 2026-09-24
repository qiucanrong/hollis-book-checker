"""Spreadsheet input/output and one controlled recovery pass."""
import csv
import codecs
import io
import logging
from collections import Counter
from copy import copy
from pathlib import Path
from itertools import islice

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.cell_range import MultiCellRange

from hollis_api import SearchError, SearchResult
from isbn_utils import normalize_isbn

LOG = logging.getLogger(__name__)
FILLS = {
    "green": PatternFill("solid", fgColor="DDF1D8"),
    "red": PatternFill("solid", fgColor="F7D9D9"),
    "yellow": PatternFill("solid", fgColor="FFF0BD"),
}
THIN = Side(style="thin", color="B7B7B7")
ALL_BORDERS = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
OUTPUT_HEADER = "Notes"
TEXT_EXTENSIONS = {".csv", ".tsv", ".txt"}


class ImportErrorMessage(ValueError):
    """A safe, user-facing input error."""


def decode_text(contents, encoding="Auto"):
    """Decode strictly: never silently replace unrecognized characters."""
    warnings = []
    if encoding != "Auto":
        chosen = encoding
    elif contents.startswith((codecs.BOM_UTF32_LE, codecs.BOM_UTF32_BE)):
        chosen = "utf-32"
    elif contents.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        chosen = "utf-16"
    else:
        # UTF-16 without a BOM needs an explicit endian choice. Trailing
        # null padding alone is not evidence of a UTF-16 encoding.
        sample = contents.rstrip(b"\x00")[:4096]
        if b"\x00" in sample:
            raise ImportErrorMessage("Select UTF-16 LE or UTF-16 BE under Text file options for this file.")
        try:
            contents.decode("utf-8-sig")
            chosen = "utf-8-sig"
        except UnicodeDecodeError:
            chosen = "cp1252"
            warnings.append("Used Windows-1252 because this file is not UTF-8. Check the preview; change the encoding if text looks wrong.")
    try:
        text = contents.decode(chosen)
    except (UnicodeError, LookupError):
        raise ImportErrorMessage("Could not decode this file. Choose another encoding under Text file options.") from None
    text = text.lstrip("\ufeff")
    padding = len(text) - len(text.rstrip("\x00"))
    if padding:
        text = text.rstrip("\x00")
        warnings.append(f"Removed {padding:,} trailing null padding characters from the text export.")
    # Embedded nulls/control characters should not silently alter book data.
    if any(ord(c) < 32 and c not in "\t\r\n" for c in text):
        raise ImportErrorMessage("The decoded file contains unsupported control characters. Check the encoding or re-export it as CSV/TSV.")
    return text, chosen, warnings


def detect_delimiter(text):
    """Compare parsed record widths, tolerating a short preamble."""
    scores = []
    for delimiter in ("\t", ",", ";", "|"):
        try:
            rows = list(islice(csv.reader(io.StringIO(text, newline=""), delimiter=delimiter), 100))
        except csv.Error:
            continue
        widths = Counter(len(row) for row in rows if len(row) > 1)
        if widths:
            width, count = widths.most_common(1)[0]
            scores.append((count, width, delimiter))
    return max(scores)[2] if scores else ","


def read_input(filename, contents, encoding="Auto", delimiter="Auto", import_info=None):
    extension = Path(filename).suffix.lower()
    if extension == ".xlsx":
        return load_workbook(io.BytesIO(contents))
    if extension in TEXT_EXTENSIONS:
        text, chosen_encoding, warnings = decode_text(contents, encoding)
        chosen_delimiter = detect_delimiter(text) if delimiter == "Auto" else delimiter
        if chosen_delimiter not in ("\t", ",", ";", "|"):
            raise ImportErrorMessage("Select a supported delimiter.")
        try:
            rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=chosen_delimiter, strict=True))
        except csv.Error:
            raise ImportErrorMessage("Could not parse the text file. Check the delimiter and quoting in the export.") from None
        if not any(any(value.strip() for value in row) for row in rows):
            raise ImportErrorMessage("This file contains no spreadsheet data.")
        if len(rows) > 1048576 or any(len(row) > 16384 for row in rows):
            raise ImportErrorMessage("This file exceeds Excel's worksheet size limits.")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Books"
        for row_index, row in enumerate(rows, 1):
            for col_index, value in enumerate(row, 1):
                if len(value) > 32767:
                    raise ImportErrorMessage("A text cell exceeds Excel's 32,767-character limit.")
                cell = sheet.cell(row_index, col_index, value)
                cell.data_type = "s"  # Preserve leading zeroes and literal '='.
        if import_info is not None:
            import_info.update(encoding=chosen_encoding, delimiter=chosen_delimiter, warnings=warnings)
        return workbook
    raise ImportErrorMessage("Upload an .xlsx, .csv, .tsv, or delimited .txt file.")


def column_options(sheet, header_row=1):
    from openpyxl.utils import get_column_letter
    options = {}
    for col in range(1, sheet.max_column + 1):
        header = sheet.cell(header_row, col).value
        if header in (None, "") and not any(sheet.cell(row, col).value not in (None, "")
                                      for row in range(header_row + 1, sheet.max_row + 1)):
            continue
        label = (str(header).strip() if header is not None else "") or "(blank header)"
        options[f"{get_column_letter(col)} — {label}"] = col
    return options


def data_rows(sheet, original_max_col, header_row=1):
    return [row for row in range(header_row + 1, sheet.max_row + 1)
            if any(sheet.cell(row, col).value not in (None, "")
                   for col in range(1, original_max_col + 1))]


def classify_rows(sheet, rows, isbn_col, client, progress=None, status=None):
    """Search each unique ISBN, then retry transient failures once.

    Persistent 429s pause the FIRST pass and queue untouched ISBNs for
    recovery. A persistent 429 in recovery ends the run. Long Retry-After
    instructions or authentication failures end it without recovery.
    """
    groups, results, outcomes = {}, {}, {}
    client.run_notice = ""
    for row in rows:
        raw = sheet.cell(row, isbn_col).value
        isbn = normalize_isbn(raw)
        if isbn is None:
            missing = raw is None or str(raw).strip() == ""
            results[row] = SearchResult("white", reason="Missing ISBN." if missing else "Invalid ISBN (format, length, or checksum).")
        else:
            groups.setdefault(isbn, []).append(row)

    def update(isbn, result):
        outcomes[isbn] = result
        for row in groups[isbn]:
            results[row] = result
        if progress:
            progress(len(results), len(rows))

    def run_pass(keys, label, recovery=False):
        for position, isbn in enumerate(keys, 1):
            if status:
                status(f"{label}: ISBN {position} of {len(keys)}.")
            error = None
            previous = outcomes.get(isbn)
            try:
                result = client.search_isbn(isbn)
            except SearchError as exc:
                error = exc
                result = SearchResult("white", reason=str(exc),
                                      retryable=exc.retryable, status_code=exc.status_code)
            except Exception as exc:
                # Log type only: exception text may contain request details.
                LOG.error("Unexpected search error: %s", type(exc).__name__)
                result = SearchResult("white", reason="Unexpected search error.")
            if recovery and previous is not None and previous.retryable and not result.retryable and error is None:
                result.reason = (result.reason + " Search completed on the recovery pass.").strip()
            update(isbn, result)
            if error is not None and error.stop_batch:
                return error
        return None

    if progress:
        progress(len(results), len(rows))
    stopped = run_pass(list(groups), "First pass")
    can_recover = stopped is None or (stopped.retryable and not stopped.defer_run)
    pending = [isbn for isbn in groups
               if isbn not in outcomes or outcomes[isbn].retryable]
    if pending and can_recover:
        try:
            client.wait_for_second_pass()
        except SearchError as exc:
            stopped = exc
        else:
            # Deliberately call the API again. Do not reuse cached failures.
            stopped = run_pass(pending, "Recovery pass", recovery=True)

    if stopped is not None:
        client.run_notice = f"Run stopped: {stopped} Completed results are available below."
    for isbn in groups:
        if isbn not in outcomes:
            reason = "Not searched because the batch was stopped."
            if stopped is not None:
                reason += " " + str(stopped)
            update(isbn, SearchResult("white", reason=reason))
        elif outcomes[isbn].retryable:
            # Keep the actual failure visible; no further automatic pass.
            if stopped is not None:
                outcomes[isbn].reason += " Further automatic retries were stopped."
    if progress:
        progress(len(rows), len(rows))
    return results


def result_notes(result):
    """Use the same notes in the preview and downloaded workbook."""
    labels = {
        "green": "No verified HOLLIS match; purchase candidate.",
        "red": "1 verified HOLLIS match.",
        "yellow": f"{len(result.matches)} verified HOLLIS matches; manual review required.",
        "white": "Manual review required.",
    }
    lines = [labels.get(result.status, "Manual review required.")]
    if result.reason:
        lines.append(result.reason)
    if result.status == "white" and result.matches:
        lines.append("Verified matches found so far (search outcome remains uncertain):")
    for index, match in enumerate(result.matches, 1):
        lines.append(f"{index}. {match.title}")
        if match.url:
            lines.append(match.url)
    return "\n".join(lines)


def render_workbook(workbook, sheet_name, rows, results, header_row=1):
    sheet = workbook[sheet_name]
    original_max_col = sheet.max_column
    merged_ranges = [copy(area) for area in sheet.merged_cells.ranges]
    dimensions = {name: copy(dim) for name, dim in sheet.column_dimensions.items()}
    sheet.insert_cols(1)  # Classification already finished using original positions.
    # openpyxl moves cells but not these layout definitions. Shift them so
    # preamble headings and existing column widths stay with their data.
    sheet.merged_cells = MultiCellRange()
    for area in merged_ranges:
        area.shift(col_shift=1)
        sheet.merged_cells.add(area)
    sheet.column_dimensions.clear()
    for name, dim in dimensions.items():
        shifted = get_column_letter(column_index_from_string(name) + 1)
        dim.index = shifted
        if dim.min:
            dim.min += 1
        if dim.max:
            dim.max += 1
        sheet.column_dimensions[shifted] = dim
    last_col = original_max_col + 1
    header = sheet.cell(header_row, 1)
    header.value = OUTPUT_HEADER
    header._style = copy(sheet.cell(header_row, 2)._style)
    sheet.column_dimensions["A"].width = 48
    counts, preview = Counter(), []
    for row in rows:
        result = results[row]
        counts[result.status] += 1
        output = sheet.cell(row, 1)
        note = result_notes(result)
        tail = "\n[Notes shortened for Excel; see the full preview.]"
        output.value = note if len(note) <= 32767 else note[:32767 - len(tail)] + tail
        output.data_type = "s"
        alignment = copy(output.alignment)
        alignment.wrap_text, alignment.vertical = True, "top"
        output.alignment = alignment
        if result.status == "red" and result.matches and result.matches[0].url:
            output.hyperlink = result.matches[0].url
            font = copy(output.font)
            font.color, font.underline = "0563C1", "single"
            output.font = font
        for col in range(1, last_col + 1):
            cell = sheet.cell(row, col)
            if result.status in FILLS:
                cell.fill = copy(FILLS[result.status])
            cell.border = copy(ALL_BORDERS)
        preview.append({"Excel row": row, "Status": result.status,
                        "Notes": note})
    for col in range(1, last_col + 1):
        sheet.cell(header_row, col).border = copy(ALL_BORDERS)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue(), counts, preview
