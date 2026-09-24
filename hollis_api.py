"""Primo API access with paced requests and batch-wide cooldowns."""
import logging
import math
import random
import re
import time
from dataclasses import dataclass, field
from datetime import timezone
from email.utils import parsedate_to_datetime
from urllib.parse import quote

import requests

from isbn_utils import equivalent, normalize_isbn

LOG = logging.getLogger(__name__)
BASE_URLS = {
    "Sandbox": "https://go.stage.apis.huit.harvard.edu/lts-exlibris-primo/primo/v1/search",
    "Production": "https://go.apis.huit.harvard.edu/lts-exlibris-primo/primo/v1/search",
}

# Tuning defaults, NOT documented Harvard quota limits.
REQUEST_GAP_SECONDS = 2.0
REQUEST_TIMEOUT = (5, 30)  # Connection timeout, read timeout.
RETRY_DELAYS = (5, 15, 30)  # Four total attempts per page per pass.
MAX_AUTO_WAIT_SECONDS = 120
SECOND_PASS_DELAY_SECONDS = 30
PAGE_SIZE = 50
MAX_RESULTS = 500
TEMPORARY_STATUSES = {429, 500, 502, 503, 504}
ISBN_TOKEN = re.compile(
    r"(?<![A-Za-z0-9])(?:\d[\s-]*){9,12}[\dXx](?![A-Za-z0-9])"
)


class SearchError(Exception):
    """Only safe, application-generated messages should be passed here."""
    def __init__(self, message, *, retryable=False, status_code=None,
                 stop_batch=False, defer_run=False):
        super().__init__(message)
        self.retryable = retryable
        self.status_code = status_code
        self.stop_batch = stop_batch
        self.defer_run = defer_run


@dataclass
class Match:
    title: str
    url: str | None = None


@dataclass
class SearchResult:
    status: str
    matches: list[Match] = field(default_factory=list)
    reason: str = ""
    retryable: bool = False
    status_code: int | None = None


def retry_after_seconds(value, now=None):
    """Support both Retry-After formats: seconds and an HTTP date."""
    if not value:
        return None
    value = str(value).strip()
    if re.fullmatch(r"[0-9]+", value):
        return int(value)
    try:
        target = parsedate_to_datetime(value)
        if target.tzinfo is None:
            target = target.replace(tzinfo=timezone.utc)
        return max(0.0, target.timestamp() - (time.time() if now is None else now))
    except (TypeError, ValueError, OverflowError):
        return None


def _first_text(value):
    if isinstance(value, list):
        value = value[0] if value else None
    return value.strip() if isinstance(value, str) else ""


def _record_isbns(record):
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
        if isinstance(value, str):
            for token in ISBN_TOKEN.findall(value):
                isbn = normalize_isbn(token)
                if isbn:
                    found.add(isbn)
    return found


def _record_title(record):
    display = record["pnx"].get("display", {})
    return _first_text(display.get("title")) if isinstance(display, dict) else ""


def _record_url(record, environment):
    if record.get("adaptor") != "Local Search Engine":
        return None
    raw_id = record.get("@id")
    if not isinstance(raw_id, str):
        return None
    identifier = raw_id.rstrip("/").split("/")[-1]
    if not re.fullmatch(r"\d+", identifier):
        return None
    host = "qa.hollis.harvard.edu" if environment == "Sandbox" else "hollis.harvard.edu"
    return f"https://{host}/permalink/01HVD_INST/1vs5jgf/alma{quote(identifier, safe='')}"


class HollisClient:
    def __init__(self, api_key, environment="Sandbox", status_callback=None):
        if environment not in BASE_URLS:
            raise ValueError("Unknown API environment.")
        self.environment = environment
        self.status_callback = status_callback
        self.session = requests.Session()
        self.session.headers.update({"X-Api-Key": api_key, "Accept": "application/json"})
        self._next_request_at = 0.0
        self._wait_reason = "Spacing out searches"
        self._deferred_error = None
        self.run_notice = ""

    def close(self):
        self.session.close()

    def _wait_until(self, deadline, message):
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return
            if self.status_callback:
                self.status_callback(f"{message}: {math.ceil(remaining)} seconds remaining.")
            time.sleep(min(1.0, remaining))

    def wait_for_second_pass(self):
        if self._deferred_error is not None:
            raise self._deferred_error
        self._wait_until(
            max(self._next_request_at, time.monotonic() + SECOND_PASS_DELAY_SECONDS),
            "Waiting before the recovery pass",
        )

    def _schedule_retry(self, attempt, header, status_code=None):
        delay = RETRY_DELAYS[min(attempt, len(RETRY_DELAYS) - 1)]
        server_delay = retry_after_seconds(header)
        # Do not shorten a server-requested delay. Stop this run instead
        # if it is too long for an interactive wait.
        if server_delay is not None and server_delay > MAX_AUTO_WAIT_SECONDS:
            error = SearchError(
                f"API requested a wait of {math.ceil(server_delay)} seconds "
                f"(HTTP {status_code}). Run stopped; retry after that interval.",
                retryable=True, status_code=status_code,
                stop_batch=True, defer_run=True,
            )
            self._deferred_error = error
            raise error
        delay = max(delay, server_delay or 0) + random.uniform(0, 1)
        self._next_request_at = max(self._next_request_at, time.monotonic() + delay)
        self._wait_reason = (
            f"API rate limit (HTTP 429); pausing all searches"
            if status_code == 429 else "Waiting before another request"
        )
        LOG.warning("env=%s HTTP=%s retry_wait=%.1fs", self.environment, status_code, delay)

    def _page(self, isbn, offset):
        if self._deferred_error is not None:
            raise self._deferred_error
        params = {
            "q": f"any,contains,{isbn}", "vid": "01HVD_INST:HVD2",
            "tab": "LibraryCatalog", "scope": "MyInstitution",
            "inst": "01HVD_INST", "offset": offset, "limit": PAGE_SIZE,
        }
        attempts = len(RETRY_DELAYS) + 1
        for attempt in range(attempts):
            self._wait_until(self._next_request_at, self._wait_reason)
            if self.status_callback:
                self.status_callback(f"Searching ISBN {isbn}, attempt {attempt + 1}/{attempts}.")
            started = time.monotonic()
            try:
                response = self.session.get(
                    BASE_URLS[self.environment], params=params,
                    timeout=REQUEST_TIMEOUT, allow_redirects=False,
                )
            except (requests.Timeout, requests.ConnectionError):
                self._next_request_at = time.monotonic() + REQUEST_GAP_SECONDS
                self._schedule_retry(attempt, None)
                if attempt == attempts - 1:
                    raise SearchError("Connection failed or timed out after retries.", retryable=True) from None
                continue
            except requests.RequestException:
                raise SearchError("Request could not be completed.") from None

            self._next_request_at = time.monotonic() + REQUEST_GAP_SECONDS
            self._wait_reason = "Spacing out searches"
            code = response.status_code
            LOG.info("env=%s isbn=%s offset=%s attempt=%s HTTP=%s elapsed=%.2fs",
                     self.environment, isbn, offset, attempt + 1, code,
                     time.monotonic() - started)
            if code in TEMPORARY_STATUSES:
                header = response.headers.get("Retry-After")
                response.close()
                self._schedule_retry(attempt, header, code)
                if attempt == attempts - 1:
                    raise SearchError(
                        f"HOLLIS request failed after retries (HTTP {code}).",
                        retryable=True, status_code=code, stop_batch=(code == 429),
                    )
                continue
            if not 200 <= code < 300:
                response.close()
                raise SearchError(
                    f"HOLLIS returned HTTP {code}.", status_code=code,
                    stop_batch=code in (401, 403),
                )
            try:
                data = response.json()
            except ValueError:
                raise SearchError("HOLLIS returned invalid JSON.") from None
            finally:
                response.close()
            if not isinstance(data, dict):
                raise SearchError("Unexpected HOLLIS response.")
            info, docs = data.get("info"), data.get("docs")
            if not isinstance(info, dict) or not isinstance(docs, list):
                raise SearchError("HOLLIS response lacks result information.")
            total = info.get("total")
            if isinstance(total, bool) or not isinstance(total, int) or total < 0:
                raise SearchError("HOLLIS response lacks a valid result count.")
            if any(not isinstance(doc, dict) for doc in docs):
                raise SearchError("Unexpected HOLLIS record format.")
            if total == 0 and docs:
                raise SearchError("HOLLIS returned inconsistent result information.")
            return total, docs
        raise SearchError("HOLLIS search failed.")

    def search_isbn(self, isbn):
        records, offset, expected_total = {}, 0, None
        while True:
            total, docs = self._page(isbn, offset)
            if expected_total is not None and total != expected_total:
                raise SearchError("Result count changed during pagination.")
            expected_total = total
            if total > MAX_RESULTS:
                raise SearchError("Too many search results to inspect safely.")
            if total == 0:
                return SearchResult("green")
            if not docs:
                raise SearchError("HOLLIS returned an incomplete result page.")
            before = len(records)
            for doc in docs:
                record_id = doc.get("@id")
                if not isinstance(record_id, str) or not record_id:
                    raise SearchError("A HOLLIS record lacks an identifier.")
                records[record_id] = doc
            if offset and len(records) == before:
                raise SearchError("HOLLIS repeated a result page.")
            offset += len(docs)
            if offset >= total:
                break

        matches, unverified = [], False
        for record in records.values():
            record_isbns = _record_isbns(record)
            if not record_isbns:
                unverified = True
                continue
            if not any(equivalent(isbn, candidate) for candidate in record_isbns):
                continue
            title = _record_title(record)
            if not title:
                unverified = True
                continue
            matches.append(Match(title, _record_url(record, self.environment)))
        LOG.info("isbn=%s candidates=%s confirmed=%s", isbn, len(records), len(matches))
        if len(matches) >= 2:
            return SearchResult("yellow", matches)
        if unverified:
            return SearchResult("white", matches, "One or more HOLLIS candidates could not be verified.")
        return SearchResult("red", matches) if matches else SearchResult("green")
