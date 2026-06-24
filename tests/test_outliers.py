"""
Test specific citations to make sure that 
the parsing logic is working correctly.
"""
import os
import pytest

from src.citation.parser import CitationParser

citations = [
    "Carson, J. V. (2019). Assessing the nuances of counterterrorism programs: A country-level investigation of",
    "Felson, R. B., Osgood, D. W., Cundiff, P. R., & Wiernik",
]

input_file = os.path.join(os.path.dirname(__file__), "data", "Journal Tracking.xlsx")
tracker = CitationParser(input_file=input_file, output_file=None)
assert tracker is not None, "Failed to create JournalTracker instance"

@pytest.mark.parametrize("citation", citations)
def test_citation_parsing(citation):
    """
    Test the parsing of a specific citation.
    """
    num_authors = tracker.num_authors(citation)
    assert num_authors > 0, f"Failed to parse authors from citation: {citation}"