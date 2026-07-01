"""
Test specific citations to make sure that 
the parsing logic is working correctly.
"""
import os
import pandas as pd
import pytest

from src.citation.parser import CitationParser


parser = CitationParser(os.path.join(os.path.dirname(__file__), '..'))
assert parser is not None, "Failed to create CitationParser instance"

test_data_location = os.path.join(os.path.dirname(__file__),
                                  'expected_output', 'expected_output.xlsx')
test_df = pd.read_excel(test_data_location, dtype=str, sheet_name="test_cases")
test_df = test_df.fillna("")
test_df = test_df.to_dict(orient="records")

@pytest.mark.parametrize("case", test_df, ids=[
    c["original_citation"][:60] for c in test_df
])
def test_regression_citation(case):
    """
    Test that the parsing of a citation matches the expected output.
    """
    original = case.pop("original_citation")
    result = parser.parse_line(original, 0, "test_sheet")
    parsed = result.to_dict()

    failures = []
    for field, expected in case.items():
        if field in ["row", "sheet_name", "parse_errors"]:
            continue  # Skip fields that are not relevant for comparison   
        actual = normalize(field, parsed.get(field, ""))
        expected = normalize(field, expected)
        if actual != expected:
            failures.append(f"  {field}: expected {expected!r}, got {actual!r}")

    assert not failures, "Mismatch in fields:\n" + "\n".join(failures)

def normalize(field, value):
    """
    Normalize a value for comparison, treating None and empty strings as equivalent.
    """
    # these fields are expected to be integers or floats, so we convert them accordingly
    if field in ["year", "parse_errors"]:
        return int(value) if value not in [None, ""] else None
    if field in ["author_num"]:
        return f'{float(value):.1f}' if value not in [None, ""] else 0
    return str(value).strip() if value not in [None, ""] else ""
