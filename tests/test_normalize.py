"""
Minimal unit tests on synthetic values — no patient data.
Run:  python -m pytest -q
These lock the normalization behaviour that underpins the concordance analysis,
including the Turkish case-folding fix described in the README.
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from validate import normalize, tr_fold


def test_numeric_formatting_equivalence():
    # decimal-formatting differences must not create false discrepancies
    assert normalize(12.0) == normalize("12")
    assert normalize(3.50) == normalize("3.5")
    assert normalize("12,4") == normalize("12.4")


def test_turkish_case_folding():
    # the artifact that produced ~480 false discrepancies must be neutralized
    assert tr_fold("ÜREDİ") == tr_fold("üredi")
    assert tr_fold("KANDİDA") == tr_fold("kandida")


def test_empty_and_none():
    assert normalize(None) == ""
    assert normalize("   ") == ""
