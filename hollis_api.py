"""Harvard APIgee Primo search and conservative response interpretation."""

import logging
import re
import time
from dataclasses import dataclass, field
from urllib.parse import quote

import requests

from isbn_utils import equivalent, normalize_isbn


LOG = logging.getLogger(__name__)

BASE_URLS = {
    "Sandbox": (
        "https://go.stage.apis.huit.harvard.edu/"
        "lts-exlibris-primo/primo/v1/search"
    ),
    "Production": (
        "https://go.apis.huit.harvard.edu/"
        "lts-exlibris-primo/primo/v1/search"
    ),
}

PAGE_SIZE = 50
MAX_RESULTS = 500
TEMPORARY_STATUSES = {429, 500, 502, 503, 504}

# Recognizes ISBN-like sequences within values such as
# "978-0-306-40615-7 (paperback)".
ISBN_TOKEN = re.compile(
    r"(?<![A-Za-z0-9])(?:\d[\s-]*){9,12}[\dXx](?![A-Za-z0-9])"
)


class SearchError(Exception):
    """A search could not be interpreted reliably."""


@dataclass
class Match:
    title: str
    url: str | None = None


@dataclass
class SearchResult:
    status: str  # green, red, yellow, white
    matches: list[Match] = field(default_factory=list)
    reason: str = ""


def _first_text(value):
    if isinstance(value, list):
        value = value[0] if value else None

    return value.strip() if isinstance(value, str) else ""


def _record_isbns(record):
    """Read ISBNs from the documented PNX addata.isbn field."""
    pnx = record.get("pnx")
    if not isinstance(pnx, dict):
        raise SearchError("A candidate record has no PNX data.")

    addata = pnx.get("addata")
    if not isinstance(addata, dict):
        return set()

    values = addata.get("isbn", [])
    if isinstance(values, str):
        values = [values]

    if not isinstance(values, list):
        raise SearchError("Unexpected ISBN field format.")

    found = set()

    for value in values:
        if not isinstance(value, str):
            continue

        for token in ISBN_TOKEN.findall(value):
            isbn = normalize_isbn(token)
            if isbn:
                found.add(isbn)

    return found


def _record_title(record):
    display = record["pnx"].get("display", {})
    if not isinstance(display, dict):
        return ""

    return _first_text(display.get("title"))


def _record_url(record, environment):
    """Construct a Harvard Primo VE local-record link."""
    if record.get("adaptor") != "Local Search Engine":
        return None

    raw_id = record.get("@id")
    if not isinstance(raw_id, str):
        return None

    identifier = raw_id.rstrip("/").split("/")[-1]
    if not re.fullmatch(r"\d+", identifier):
        return None

    host = (
        "qa.hollis.harvard.edu"
        if environment == "Sandbox"
        else "hollis.harvard.edu"
    )

    return (
        f"https://{host}/permalink/01HVD_INST/1vs5jgf/"
        f"alma{quote(identifier, safe='')}"
    )


class HollisClient:
    def __init__(self, api_key, environment="Sandbox"):
        if environment not in BASE_URLS:
            raise ValueError("Unknown API environment.")

        self.environment = environment
        self.session = requests.Session()
        self.session.headers.update({
            "X-Api-Key": api_key,
            "Accept": "application/json",
        })

    def close(self):
        self.session.close()

    def _page(self, isbn, offset):
        params = {
            "q": f"any,contains,{isbn}",
            "vid": "01HVD_INST:HVD2",
            "tab": "LibraryCatalog",
            "scope": "MyInstitution",
            "inst": "01HVD_INST",
            "offset": offset,
            "limit": PAGE_SIZE,
        }

        for attempt in range(3):
            try:
                response = self.session.get(
                    BASE_URLS[self.environment],
                    params=params,
                    timeout=(5, 25),
                )
            except requests.RequestException:
                if attempt == 2:
                    raise SearchError(
                        "HOLLIS connection failed."
                    ) from None

                time.sleep(1 + attempt)
                continue

            LOG.info(
                "Primo HTTP status: %s",
                response.status_code,
            )

            if response.status_code in TEMPORARY_STATUSES:
                if attempt == 2:
                    raise SearchError(
                        f"HOLLIS request failed after retries "
                        f"(HTTP {response.status_code}). Please retry later."
                    )

                retry_after = response.headers.get("Retry-After", "")
                delay = (
                    min(int(retry_after), 8)
                    if retry_after.isdigit()
                    else 1 + attempt
                )
                time.sleep(delay)
                continue

            if not response.ok:
                raise SearchError(
                    f"HOLLIS returned HTTP {response.status_code}."
                )

            try:
                data = response.json()
            except ValueError:
                raise SearchError(
                    "HOLLIS returned invalid JSON."
                ) from None

            if not isinstance(data, dict):
                raise SearchError("Unexpected HOLLIS response.")

            info = data.get("info")
            docs = data.get("docs")

            if not isinstance(info, dict) or not isinstance(docs, list):
                raise SearchError(
                    "HOLLIS response lacks result information."
                )

            total = info.get("total")
            if (
                isinstance(total, bool)
                or not isinstance(total, int)
                or total < 0
            ):
                raise SearchError(
                    "HOLLIS response lacks a valid result count."
                )

            if any(not isinstance(doc, dict) for doc in docs):
                raise SearchError(
                    "Unexpected HOLLIS record format."
                )

            return total, docs

        raise SearchError("HOLLIS search failed.")

    def search_isbn(self, isbn):
        """Return white if candidates or pages prevent a safe decision."""
        records = {}
        offset = 0

        while True:
            total, docs = self._page(isbn, offset)

            if total > MAX_RESULTS:
                raise SearchError(
                    "Too many search results to inspect safely."
                )

            if total == 0:
                return SearchResult("green")

            if not docs:
                raise SearchError(
                    "HOLLIS returned an incomplete result page."
                )

            for doc in docs:
                record_id = doc.get("@id")

                if not isinstance(record_id, str) or not record_id:
                    raise SearchError(
                        "A HOLLIS record lacks an identifier."
                    )

                records[record_id] = doc

            offset += len(docs)

            if offset >= total:
                break

            if offset > MAX_RESULTS:
                raise SearchError("Too many search results.")

            time.sleep(0.2)

        matches = []
        unverified = False

        for record in records.values():
            record_isbns = _record_isbns(record)

            if not record_isbns:
                unverified = True
                continue

            if not any(
                equivalent(isbn, candidate)
                for candidate in record_isbns
            ):
                continue

            title = _record_title(record)

            if not title:
                unverified = True
                continue

            matches.append(
                Match(
                    title=title,
                    url=_record_url(record, self.environment),
                )
            )

        LOG.info(
            "ISBN %s: %s candidates, %s confirmed",
            isbn,
            len(records),
            len(matches),
        )

        if len(matches) >= 2:
            return SearchResult("yellow", matches)

        if unverified:
            return SearchResult(
                "white",
                matches,
                "One or more HOLLIS candidates could not be verified.",
            )

        if len(matches) == 1:
            return SearchResult("red", matches)

        return SearchResult("green")