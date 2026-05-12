"""Tests for the title_case helper used to fix ALL-CAPS proper-noun rendering.

AM caught Tao Dang's notice rendering 'TAO DANG' on 2026-05-11 — the upstream
Airtable record stores names in all caps but the notice belongs to "Tao Dang".
We need a defensive title-case helper that fixes obvious all-caps input
without mangling intentional capitalization like McDonald or O'Brien.
"""
import os
import sys

os.environ.setdefault("DROPBOX_SIGN_API_KEY", "dummy")
os.environ.setdefault("AIRTABLE_API_KEY", "dummy")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), os.pardir))

from main import title_case  # noqa: E402


def test_empty_string():
    assert title_case("") == ""


def test_none_returns_empty():
    assert title_case(None) == ""


def test_all_caps_simple_name():
    assert title_case("TAO DANG") == "Tao Dang"


def test_all_lower_simple_name():
    assert title_case("tao dang") == "Tao Dang"


def test_already_proper_case_preserved():
    assert title_case("Tao Dang") == "Tao Dang"


def test_all_caps_with_apostrophe():
    assert title_case("O'BRIEN") == "O'Brien"


def test_all_caps_with_hyphen():
    assert title_case("MARY-ANNE") == "Mary-Anne"


def test_mixed_proper_with_hyphen_preserved():
    assert title_case("Anne-Michelle") == "Anne-Michelle"


def test_all_caps_with_initials():
    assert title_case("J.R. SMITH") == "J.R. Smith"


def test_mixed_case_preserved_mcdonald():
    # Intentional internal capital — don't mangle.
    assert title_case("McDonald") == "McDonald"


def test_all_caps_street_address():
    assert title_case("123 MAIN ST") == "123 Main St"


def test_whitespace_trimmed():
    assert title_case("  TAO DANG  ") == "Tao Dang"


def test_blank_line_sentinel_preserved():
    # The handler uses '_______________' as a sentinel for missing values.
    # title_case must not mangle it.
    assert title_case("_______________") == "_______________"
