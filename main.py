"""
Run this script to perform the work.
"""
import sys

import pytest

from src.citation.parser import CitationParser

def main():
    """
    Main function to run the journal tracker.
    """

    print("Welcome to the Journal Tracker!")
    print("This script will parse journal citations from an Excel file and output structured data.")
    print("Please ensure that the input Excel file is located in the 'data' directory.")
    print("The output will be saved in the 'data' directory as well.")
    print("First we will run the tests for our outliers and then the parser if those all pass.")

    result = pytest.main(["tests/", "-q", "--tb=short"])
    if result != pytest.ExitCode.OK:
        print("Tests failed — aborting.", file=sys.stderr)
        sys.exit(1)

    parser = CitationParser()
    parser.run()


if __name__ == "__main__":
    main()
