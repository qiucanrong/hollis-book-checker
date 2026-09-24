# HOLLIS ISBN Purchase Screening

A Streamlit application for checking book lists against the Harvard
Library HOLLIS catalog through Harvard APIgee's Primo API.

## Local setup

Use Python 3.10 or newer.

Create and activate a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Install dependencies:

```bash
python -m pip install -r requirements.txt
```

Create `.streamlit/secrets.toml` containing:

```toml
HOLLIS_API_KEY = "your-actual-api-key"
```

Keep this file out of GitHub. The supplied `.gitignore` excludes it.

Run the application from the project folder:

```bash
python -m streamlit run app.py
```

The API key must correspond to the environment selected in the app.
The API secret is not used.

## Input

Supported formats:

- Excel `.xlsx`
- UTF-8 `.csv`

The first row is treated as the header.

Select the worksheet, ISBN column, and original title column.
Each ISBN cell should contain one ISBN.

The app removes spaces and hyphens, handles integer-valued Excel
numbers, and validates ISBN-10 or ISBN-13 length and checksum.

Missing or invalid ISBNs remain uncolored for manual review.
Cells containing multiple identifiers require manual review.

## Search behavior

The app searches the normalized ISBN using:

`q=any,contains,<ISBN>`

The search is limited to HOLLIS Library Catalog.

Returned candidates are checked against their `pnx.addata.isbn`
values. Equivalent ISBN-10 and ISBN-13 values count as matches.
Records are deduplicated by their returned identifiers.

Rows are classified as:

- Green: no confirmed matching record.
- Red: exactly one confirmed matching record.
- Yellow: multiple confirmed matching records.
- Uncolored: missing/invalid input, search error, or insufficient
  information to determine the match count reliably.

Repeated ISBNs reuse their results within a processing run.

## Output

The downloaded filename is:

`Searched_original_filename.xlsx`

`Title of Book Found` is inserted in column A of the selected sheet.
Original columns move one position to the right.

The app colors the processed rows and adds borders to the header
and processed cells.

A red row's found title links to HOLLIS when a local-record link
can be constructed. Yellow rows list matching titles and available
URLs in the same cell.

Review notes appear in the app's results preview.

Other worksheets remain in the output workbook.

Column insertion does not automatically repair existing formula
references, merged ranges, table references, or column widths.
This output approach is intended for ordinary tabular book lists.

## Streamlit Community Cloud

Upload the source files and requirements.txt to GitHub.
Do not upload `.streamlit/secrets.toml`.

Select `app.py` as the application's entry point.

In the app's Streamlit Cloud secrets settings, enter:

```toml
HOLLIS_API_KEY = "your-actual-api-key"
```

## Logging

Logging defaults to WARNING.

For more diagnostic information, set `HOLLIS_LOG_LEVEL` to `INFO`
in the server environment. Application logging includes HTTP
statuses and search counts, without logging the API key.

## References

- Harvard Primo APIgee instructions:
  https://harvardwiki.atlassian.net/wiki/spaces/LibraryStaffDoc/pages/896663553

- Harvard Primo VE query and result guidance:
  https://harvardwiki.atlassian.net/wiki/spaces/LibraryStaffDoc/pages/43422622

- Ex Libris Primo response documentation:
  https://developers.exlibrisgroup.com/primo/apis/search-output/