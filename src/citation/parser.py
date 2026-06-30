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
        if not line or not isinstance(line, str):
            self.logger.warning(
                'Empty or invalid citation line at row %s in sheet "%s".',
                row, sheet_name)
            return Citation(original_citation=line, row=row, sheet_name=sheet_name, logger=self.logger)

        citation = Citation(original_citation=line, row=row, sheet_name=sheet_name, logger=self.logger)
        # Detect year in parentheses
        year_match = re.search(r'\((\d{4})\)', citation.original_citation)
        citation.year = year_match.group(1) if year_match else ''

        # Extract DOI or URL
        doi_match = re.search(r'(https?://\S+)', citation.original_citation)
        if doi_match:
            citation.doi = doi_match.group(1)

        if citation.year:
            # Split authors vs. rest of citation
            parts = citation.original_citation.split(f'({citation.year})')
            citation.authors = parts[0].strip().rstrip('.')
            rest = parts[1]

            # Remove DOI portion from the tail before parsing journal info
            if citation.doi:
                rest = rest.split(citation.doi)[0]

            rest = rest.replace(" [page numbers missing]", "")
            rest_parts = rest.split(',')
            if citation.doi == 'https://doi.org/10.1007/s10940-016-9332-7':
                print(f'rest: {rest}, rest_parts: {rest_parts}')

            pages_match = re.search(r'(\d+[-–][\d\?]+)\.*\s*$', rest_parts[-1].strip())
            citation.pages = rest_parts.pop(-1).strip().rstrip('.') if len(rest_parts) > 1 and pages_match else ''

            volissue_match = re.search(r'(XX|\d+(\([-–\d]+\))*)\s*$', rest_parts[-1].strip())
            citation.volume_issue = rest_parts.pop(-1).strip() if len(rest_parts) > 1 and volissue_match else ''

            title_journal = ','.join(rest_parts).lstrip('.').strip() if len(rest_parts) > 0 else ''
            title_match = re.match(r'(.*[\.\?\!])([^\.\?\!]+)$', title_journal)
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
                print(f'{sheet_name}')
                df = xl.parse(sheet_name=sheet_name)

                entries = []
                for i, row in df.iterrows():
                    line = (
                        row['Full Citation']
                        if sheet_name == 'british j of crim'
                        else row['Citation'])

                    citation = self.parse_line(line, i, sheet_name)
                    citation.validate()
                    entries.append(citation.to_dict())

                # Convert to DataFrame and export
                df = pd.DataFrame(entries)
                df.to_excel(writer, sheet_name=str(sheet_name), index=False, engine='openpyxl')

                if len(entries) == 0:
                    warning_count = 0
                    self.logger.warning('sheet "%s" has no entries', sheet_name)
                else:
                    warning_count = (df['parse_errors']).sum()
                output_message = f'created sheet "{sheet_name}" with {warning_count} mis-parsed citations.'
                if warning_count > 0:
                    self.logger.error(output_message)
                else:
                    self.logger.info(output_message)

        self.logger.info('\n✅ Finished parsing citations. Output written to %s\n',
                         self.output_file)
