"""
Test specific citations to make sure that 
the parsing logic is working correctly.
"""
import os

import pytest

from src.citation.parser import CitationParser

citations = [
    "Gadd, D. (2015). In the aftermath of violence: What constitutes a responsive response? British Journal of Criminology, 55(6), 1031–1039. https://doi.org/10.1093/bjc/azv096",
    "Crow, M. S., & Goulette, N. (2025). U.S. district court judicial diversity: The impact of race and sex composition on sentencing outcomes at the district level. Crime & Delinquency, 71(11), 3529–3553. https://doi.org/10.1177/00111287241231748",
    "Boehme, H. M., Williams, T. V., Brown, N., Kidd, L., Hernandez, B., & Nolan, M. S. (2022). “It’s all about just creating the safe space”: Barbershops and beauty salons as community anchors in Black neighborhoods: Crime prevention, cohesion, and support during the COVID-19 pandemic. Crime & Delinquency, 1–25. https://doi.org/10.1177/00111287221130956",
    "Levin, A., Rosenfeld, R., & Deckard, M. (2017). The law of crime concentration: An application and recommendations for future research. Journal of Quantitative Criminology, 33(3), 633–? [page numbers missing]. https://doi.org/10.1007/s10940-016-9332-7"
]
expected_failures = [
    "Carson, J. V. (2019). Assessing the nuances of counterterrorism programs: A country-level investigation of",
    "Felson, R. B., Osgood, D. W., Cundiff, P. R., & Wiernik"
]

tracker = CitationParser(os.path.join(os.path.dirname(__file__), '..'))
assert tracker is not None, "Failed to create JournalTracker instance"

@pytest.mark.parametrize("citation", citations)
def test_citation_parsing(citation):
    """
    Test the parsing of a specific citation.
    """
    parsed = tracker.parse_line(citation, 0, "test_sheet")
    print(f"Testing citation: {citation}, Parsed: {vars(parsed)}")
    for key, value in vars(parsed).items():
        assert value is not None, f"Failed to parse {key} from citation: {citation}"
