"""Normalize and validate ISBN-10 and ISBN-13 values."""

import math
import re


def normalize_isbn(value):
    """Return a validated, unhyphenated ISBN, or None."""
    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer():
            return None
        value = str(int(value))
    else:
        value = str(value).strip()
        if value.endswith(".0") and value[:-2].isdigit():
            value = value[:-2]

    value = re.sub(r"[\s-]", "", value).upper()

    if len(value) == 10:
        if not re.fullmatch(r"\d{9}[\dX]", value):
            return None

        total = sum(
            (10 - i) * int(c)
            for i, c in enumerate(value[:9])
        )
        total += 10 if value[-1] == "X" else int(value[-1])

        return value if total % 11 == 0 else None

    if len(value) == 13 and value.isdigit():
        total = sum(
            int(c) * (1 if i % 2 == 0 else 3)
            for i, c in enumerate(value)
        )
        return value if total % 10 == 0 else None

    return None


def isbn13(isbn):
    """Convert a validated ISBN-10 to ISBN-13 for comparisons."""
    if len(isbn) == 13:
        return isbn

    base = "978" + isbn[:9]
    total = sum(
        int(c) * (1 if i % 2 == 0 else 3)
        for i, c in enumerate(base)
    )
    return base + str((-total) % 10)


def equivalent(a, b):
    """Compare validated ISBNs, allowing ISBN-10/ISBN-13 equivalence."""
    return isbn13(a) == isbn13(b)