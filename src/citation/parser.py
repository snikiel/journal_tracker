"""
Takes an Excel file of journal citations and parses them 
into a new Excel file with structured fields.
"""
import argparse
import logging
import os
import re
import sys
import time

import pandas as pd
from termcolor import colored

from src.citation.citation import Citation

class ColorFormatter(logging.Formatter):
    """
    Adds clarity in stdout by colorizing log messages based on their level.
    """
    COLORS = {
        logging.DEBUG:    "white",
        logging.INFO:     "green",
        logging.WARNING:  "yellow",
        logging.ERROR:    "red",
        logging.CRITICAL: "red",
    }

    def format(self, record: logging.LogRecord) -> str:
        msg = super().format(record)
        if not sys.stdout.isatty():
            return msg
        return colored(msg, self.COLORS.get(record.levelno, "white"))

class CitationParser():
    """
    Class to parse journal citations from an Excel file and output structured data.
    """
    def __init__(self, root=os.path.dirname(os.path.abspath(sys.argv[0]))):
        args = self.get_args()
        self.root = root
        self.input_file = os.path.join(self.root, "data", args.input)
        self.output_file = os.path.join(self.root, "data", args.output)
        self.warning_count = {}
        self.logger = self.init_logger()

    def init_logger(self):
        """
        Initialize the logger for the CitationParser.
        """
        fmt = '%(asctime)s - %(levelname)s - %(message)s'

        file_handler = logging.FileHandler(
            os.path.join(self.root, "logs", f"citation_parser-{round(time.time())}.log"),
            encoding='utf-8'
        )
        file_handler.setFormatter(logging.Formatter(fmt))

        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setFormatter(ColorFormatter(fmt))

        logging.basicConfig(
            level=logging.DEBUG,
            handlers=[file_handler, stream_handler]
        )
        return logging.getLogger(__name__)

    def get_args(self):
        """
        Get command-line arguments.
        """
        parser = argparse.ArgumentParser(
            prog="journal_tracker",
            description="Parse journal citations from an Excel file.")
        parser.add_argument(
            "-i","--input",
            type=str,
            default="Journal Tracking.xlsx",
            help="Path to the input Excel file containing journal citations.")
        parser.add_argument(
            "-o","--output",
            type=str,
            default="Journal Tracking - parsed.xlsx",
            help="Path to the output Excel file to write parsed citations.")

        args = parser.parse_args()
        return args


    def parse_line(self, line, row, sheet_name):
        """
        Parse a single citation line into structured fields.
        """
        citation = Citation(
            original_citation=line, row=row, sheet_name=sheet_name, logger=self.logger)

        if not line or not isinstance(line, str):
            self.logger.warning(
                'Empty or invalid citation line at row %s in sheet "%s".',
                row, sheet_name)
            return citation

        # Detect year in parentheses
        year_match = re.search(r'(.*)\((\d{4})\)(.*)', citation.original_citation)
        if year_match:
            citation.authors = year_match.group(1).strip().rstrip('.')
            citation.year = year_match.group(2) if year_match else ''
            rest = year_match.group(3)

            # Extract DOI or URL
            doi_match = re.search(r'(.*)(https?://\S+)', rest)
            if doi_match:
                rest = doi_match.group(1)
                citation.doi = doi_match.group(2)

            # Some entries include this text in addition to a question mark next to the page nums.
            rest = rest.replace(" [page numbers missing]", "")
            rest_parts = rest.split(',')

            # Page listings can be 1 or more pages separated by a dash
            # and are often followed by a period.
            pages_match = re.search(r'(\d+[-–][\d\?]+)\.*\s*$', rest_parts[-1].strip())
            citation.pages = rest_parts.pop(-1).strip().rstrip('.') if len(rest_parts) > 1 and pages_match else ''

            # Volumes and issues are tethered in a Volume(Issue) format.
            # Some entries include muleiple issues, and some missing volumes are entered as XX.
            volissue_match = re.search(r'(XX|\d+(\([-–\d]+\))*)\s*$', rest_parts[-1].strip())
            citation.volume_issue = rest_parts.pop(-1).strip() if len(rest_parts) > 1 and volissue_match else ''

            # What's left at this point should be just the title and the journal.
            # None of the journals have any punctuation in them, and each of the articles seem
            # to end with one, so we're splitting on that.
            title_journal = ','.join(rest_parts).lstrip('.').strip() if len(rest_parts) > 0 else ''
            title_match = re.match(r'(.*[”\.\?\!])([^”\.\?\!]+)$', title_journal)
            if title_match:
                citation.title = title_match.group(1).strip().rstrip('.')
                citation.journal = title_match.group(2).strip().rstrip('.')

        else:
            # No year present = author-only entry
            citation.authors = citation.original_citation.strip().rstrip('.')
            self.logger.warning(
                'No year found in citation: "%s"(row#%s in sheet "%s")'
                , citation.original_citation, row, sheet_name)

        citation.num_authors()
        return citation

    def run(self):
        """
        Main function to read the input Excel file, parse citations, and write to output Excel file.
        """

        # Read source file
        try:
            xl = pd.ExcelFile(self.input_file)
        except FileNotFoundError:
            self.logger.error("Input file not found: %s", self.input_file)
            self.logger.error(
                "Exiting without processing. " +
                "Please ensure the input file exists in the 'data' directory.")
            sys.exit(1)

        # Open new file for writing, so we can loop through all sheets
        # without knowing/caring what they're called.
        with pd.ExcelWriter(self.output_file) as writer:
            for sheet_name in xl.sheet_names:
                self.logger.debug('Starting the parsing of sheet "%s"', sheet_name)
                df = xl.parse(sheet_name=sheet_name)

                entries = []
                for i, row in df.iterrows():
                    line = (
                        row['Full Citation']
                        if sheet_name == 'british j of crim'
                        else row['Citation'])

                    line_number = int(str(i)) + 2
                    citation = self.parse_line(line, line_number, sheet_name)
                    citation.validate()
                    entries.append(citation.to_dict())

                # Convert to DataFrame and export
                df = pd.DataFrame(entries)
                df.to_excel(writer, sheet_name=str(sheet_name), index=False, engine='openpyxl')

                if len(entries) == 0:
                    warning_count = 0
                    self.logger.warning('sheet "%s" has no entries', sheet_name)
                else:
                    journal_names = df['journal'].unique()
                    if len(journal_names) > 1:
                        self.logger.warning('sheet %s has multiple journal names: %s',
                                            sheet_name, journal_names)
                    warning_count = (df['parse_errors']).sum()
                output_message = f'created sheet "{sheet_name}" with {warning_count} mis-parsed citations.'
                if warning_count > 0:
                    self.logger.error(output_message)
                else:
                    self.logger.info(output_message)

        self.logger.info('\n✅ Finished parsing citations. Output written to %s\n',
                         self.output_file)
