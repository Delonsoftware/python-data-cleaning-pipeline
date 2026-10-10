"""
Tests for the cleaning pipeline.

Run with either:
    python test_clean_data.py
    pytest -q          (if you have pytest installed)
"""

import re
import tempfile
from pathlib import Path

import pandas as pd

import clean_data as cd
import generate_messy_data as gen


def test_names():
    assert cd.clean_name("SMITH, JOHN")[0] == "John Smith"
    assert cd.clean_name("  mary   o'brien ")[0] == "Mary O'Brien"
    assert cd.clean_name("ronald mcdonald")[0] == "Ronald McDonald"
    assert cd.clean_name("jean-luc picard7")[0] == "Jean-Luc Picard"
    assert cd.clean_name("1234")[0] is None


def test_emails():
    assert cd.clean_email(" Jane.Doe@GMAIL.com ")[0] == "jane.doe@gmail.com"
    assert cd.clean_email("jane@gmial.com")[0] == "jane@gmail.com"
    assert cd.clean_email("jane.doe.gmail.com")[0] is None
    assert cd.clean_email("a@b")[0] is None
    assert cd.clean_email("a..b@gmail.com")[0] is None


def test_phones():
    for raw in ["416-555-0187", "(416) 555-0187", "416.555.0187", "4165550187",
                "+1 416 555 0187", "1-416-555-0187", "416 555 0187 ext 22"]:
        assert cd.clean_phone(raw)[0] == "(416) 555-0187", raw
    for bad in ["555-0187", "123", "000-000-0000", "416-555"]:
        assert cd.clean_phone(bad)[0] is None, bad


def test_regions():
    for raw in ["ON", "on", "Ontario", "ontario", "Ont."]:
        assert cd.clean_region(raw)[0] == "ON", raw
    assert cd.clean_region("new york")[0] == "NY"
    assert cd.clean_region("Québec")[0] == "QC"
    assert cd.clean_region("Narnia")[0] is None


def test_dates():
    assert cd.clean_date("2024-03-05")[0] == "2024-03-05"
    assert cd.clean_date("03/05/2024")[0] == "2024-03-05"      # month first
    assert cd.clean_date("Mar 05, 2024")[0] == "2024-03-05"
    assert cd.clean_date("05-Mar-2024")[0] == "2024-03-05"
    assert cd.clean_date("March 5, 2024")[0] == "2024-03-05"
    assert cd.clean_date("not sure")[0] is None
    assert cd.clean_date("32/13/2024")[0] is None
    assert cd.clean_date("2099-01-01")[0] is None               # future


def test_amounts():
    assert cd.clean_spend("$1,234.50")[0] == 1234.5
    assert cd.clean_spend("CAD 120")[0] == 120.0
    assert cd.clean_spend("-50")[0] is None
    assert cd.clean_spend("abc")[0] is None


def test_placeholders_are_blank():
    for token in ["", "N/A", "null", "None", "-", "  "]:
        assert cd.is_missing(token)
    assert not cd.is_missing("John")


def test_shared_email_with_different_names_is_not_merged():
    df = pd.DataFrame({
        "customer_id": ["C1", "C2", "C3"],
        "full_name": ["Anita Robinson", "Aisha Robinson", "Anita Robinson"],
        "email": ["shared@gmail.com"] * 3,
        "phone": [None, None, None],
    })
    conflicts = []
    groups = cd.group_ids(df, conflicts)
    assert groups[0] == groups[2]          # same person -> merged
    assert groups[0] != groups[1]          # different person -> kept apart
    assert len(conflicts) == 1             # ...and flagged for review


def test_end_to_end_pipeline():
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        rows = gen.generate(rows=500, seed=7)
        messy = pd.DataFrame(rows, columns=gen.FIELDS)
        messy.to_csv(tmp / "messy.csv", index=False)

        result = cd.run_pipeline(tmp / "messy.csv", tmp / "out")
        clean = pd.read_csv(tmp / "out" / "clean_customers.csv", dtype=str)

        # fewer rows out than in, and the books balance
        assert result["output_rows"] < result["input_rows"]
        assert result["input_rows"] == (result["output_rows"]
                                        + result["duplicates_removed"]
                                        + result["rejected"])

        # every remaining value is in standard form
        assert clean["email"].dropna().str.match(cd.EMAIL_RE).all()
        assert clean["email"].dropna().is_unique
        assert clean["phone"].dropna().str.fullmatch(r"\(\d{3}\) \d{3}-\d{4}").all()
        assert clean["province_state"].dropna().str.fullmatch(r"[A-Z]{2}").all()
        assert clean["signup_date"].dropna().str.fullmatch(r"\d{4}-\d{2}-\d{2}").all()
        assert (clean["total_spend"].dropna().astype(float) >= 0).all()

        # nothing without contact info survives
        assert (clean["email"].notna() | clean["phone"].notna()).all()

        # all four output files exist
        for name in ["clean_customers.csv", "change_log.csv",
                     "rejected_rows.csv", "cleaning_report.xlsx"]:
            assert (tmp / "out" / name).exists(), name

        # running the cleaner on its own output changes nothing (idempotent)
        second = cd.run_pipeline(tmp / "out" / "clean_customers.csv", tmp / "out2")
        assert second["output_rows"] == result["output_rows"]
        assert len(second["changes"]) == 0


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith("test_") and callable(f)]
    failed = 0
    for name, fn in tests:
        try:
            fn()
            print(f"PASS  {name}")
        except AssertionError as exc:
            failed += 1
            print(f"FAIL  {name}  {exc}")
    print(f"\n{len(tests) - failed}/{len(tests)} tests passed")
    raise SystemExit(1 if failed else 0)
