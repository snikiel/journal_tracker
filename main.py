"""
Run this script to perform the work.
"""
import argparse
import os

from src.citation.parser import CitationParser

def get_args():
    """
    Get command-line arguments.
    """
    parser = argparse.ArgumentParser(
        prog="journal_tracker",
        description="Parse journal citations from an Excel file.")
    parser.add_argument(
        "-if","--input",
        type=str,
        default="Journal Tracking.xlsx",
        help="Path to the input Excel file containing journal citations.")
    parser.add_argument(
        "-of","--output",
        type=str,
        default="parsed_full.xlsx",
        help="Path to the output Excel file to write parsed citations.")

    args = parser.parse_args()
    return args

def main():
    """
    Main function to run the journal tracker.
    """

    print("Hello from journal-tracker!")
    args = get_args()
    input_file = os.path.join(os.path.dirname(__file__), "data", args.input)
    output_file = os.path.join(os.path.dirname(__file__), "data", args.output)
    tracker = CitationParser(input_file=input_file, output_file=output_file)
    tracker.run()


if __name__ == "__main__":
    main()
