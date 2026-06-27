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
    "Crow, M. S., & Goulette, N. (2025). U.S. district court judicial diversity: The impact of race and sex composition on sentencing outcomes at the district level. Crime & Delinquency, 71(11), 3529–3553. https://doi.org/10.1177/00111287241231748"
]

input_file = os.path.join(os.path.dirname(__file__), "data", "Journal Tracking.xlsx")
tracker = CitationParser()
assert tracker is not None, "Failed to create JournalTracker instance"

@pytest.mark.parametrize("citation", citations)
def test_citation_parsing(citation):
    """
    Test the parsing of a specific citation.
    """
    num_authors = tracker.num_authors(citation)
    assert num_authors > 0, f"Failed to parse authors from citation: {citation}"