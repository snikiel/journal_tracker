"""
Takes an Excel file of journal citations and parses them 
into a new Excel file with structured fields.
"""
import argparse
import logging
import os
import pprint
import re
import sys
from termcolor import colored
import time
from numpy import floor
import pandas as pd

class CitationParser():
    """
    Class to parse journal citations from an Excel file and output structured data.
    """
    def __init__(self):
        args = self.get_args()
        self.root = os.path.dirname(os.path.abspath(sys.argv[0]))
        self.input_file = os.path.join(self.root, "data", args.input)
        self.output_file = os.path.join(self.root, "data", args.output)
        self.warning_count = {}
        self.logger = self.init_logger()

    def init_logger(self):
        """
        Initialize the logger for the CitationParser.
        """
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(os.path.join(self.root, "logs", f"citation_parser-{round(time.time())}.log"), encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
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
            default="parsed_full.xlsx",
            help="Path to the output Excel file to write parsed citations.")

        args = parser.parse_args()
        return args

    def num_authors(self, s):
        """Takes a string of authors and returns the number of authors"""
        omit_suffixes = ["Jr", "Sr", "Jr.", "Sr.", "III", "IV"]
        names = s.split(',')
        filtered_names = [name for name in names if name.strip() not in omit_suffixes]
        res = len(filtered_names)/2

        if not res.is_integer():
            print(f'"{s}" counted {res} authors')
        return res

    def run(self):
        """
        Main function to read the input Excel file, parse citations, and write to output Excel file.
        """
        print("Hello from journal-tracker!")

        # Read source file
        xl = pd.ExcelFile(self.input_file)

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

                    # Detect year in parentheses
                    year_match = re.search(r'\((\d{4})\)', line)
                    year = year_match.group(1) if year_match else ''

                    # Extract DOI or URL
                    doi = ''
                    doi_match = re.search(r'(https?://\S+)', line)
                    if doi_match:
                        doi = doi_match.group(1)

                    # Prepare default fields
                    title = ''
                    journal = ''
                    volissue = ''
                    pages = ''

                    if year:
                        # Split authors vs. rest of citation
                        parts = line.split(f'({year})')
                        authors = parts[0].strip().rstrip('.')
                        rest = parts[1]

                        # Match title, then everything after
                        title_match = re.match(r'\.\s*(.*?)\.\s*(.*)', rest)
                        if title_match:
                            title = title_match.group(1).strip()
                            tail = title_match.group(2)

                            # Remove DOI portion from the tail before parsing journal info
                            if doi:
                                tail = tail.split(doi)[0]

                            # Try to capture Journal, Volume(Issue), Pages
                            journal_match = re.match(
                                r'(.*?),\s*(\d+\([^)]*\)),\s*([\d–-]+)',
                                tail
                            )

                            if journal_match:
                                journal = journal_match.group(1).strip()
                                volissue = journal_match.group(2)
                                pages = journal_match.group(3)

                    else:
                        # No year present = author-only entry
                        authors = line
                        print(f'No year found in citation: "{line}"(row#{i} in sheet "{sheet_name}")')


                    citation = {
                        'Original Citation': line,
                        'Authors': authors,
                        'Year': year,
                        'Title': title,
                        'Journal': journal,
                        'VolumeIssue': volissue,
                        'Pages': pages,
                        'DOI': doi,
                        'Number of Authors': self.num_authors(authors),
                    }
                    self.validate_citation(citation, i, sheet_name)
                    entries.append(citation)

                # Convert to DataFrame and export
                df = pd.DataFrame(entries)
                df.to_excel(writer, sheet_name=str(sheet_name), index=False, engine='openpyxl')

                print(f'created sheet {sheet_name} with {self.output_warning_count(sheet_name)} parsing warnings')
                exit()

    def validate_citation(self, citation, row_index, sheet_name):
        """
        Validate the parsed citation fields and print warnings for any issues.
        """
        if not citation['Year']:
            self.log_parsing_warning("No year found", citation, row_index, sheet_name)
        if not citation['Title']:
            self.log_parsing_warning("No title found", citation, row_index, sheet_name)
        if not citation['Journal']:
            self.log_parsing_warning("No journal found", citation, row_index, sheet_name)
        if citation['Journal'] and len(citation['Journal']) > 0 and citation['Journal'][0] != citation['Journal'][0].upper():
            self.log_parsing_warning("Journal name doesn't match sheet name", citation, row_index, sheet_name)
        if not citation['VolumeIssue']:
            self.log_parsing_warning("No volume/issue found", citation, row_index, sheet_name)
        if not citation['Pages']:
            self.log_parsing_warning("No pages found", citation, row_index, sheet_name)
        if not citation['DOI']:
            self.log_parsing_warning("No DOI found", citation, row_index, sheet_name)
        if citation['Number of Authors'] != floor(citation['Number of Authors']):
            self.log_parsing_warning("Number of authors is not an integer", citation, row_index, sheet_name)

    def log_parsing_warning(self, message, citation, row_index, sheet_name):
        """
        Log a parsing warning with details about the citation and its location.
        """
        self.logger.warning(
            'Warning: (sheet: "%s", row: %s) - %s in citation: %s', 
            sheet_name, row_index, message, citation)
        print(colored('Warning:', 'yellow') + 
              f' (sheet: "{sheet_name}", row: {row_index}) - {message} in citation:')
        pprint.pprint(citation)
        self.tally_parsing_warnings(citation, sheet_name)

    def tally_parsing_warnings(self, citation, sheet_name):
        """
        Tally the number of parsing warnings for each citation.
        """
        if sheet_name not in self.warning_count:
            self.warning_count[sheet_name] = {}
        if citation['Original Citation'] not in self.warning_count[sheet_name]:
            self.warning_count[sheet_name][citation['Original Citation']] = 1
        else:
            self.warning_count[sheet_name][citation['Original Citation']] += 1

    def output_warning_count(self, sheet_name):
        """
        Output the count of parsing warnings for each citation in the specified sheet.
        """
        if sheet_name in self.warning_count:
            return len(self.warning_count[sheet_name])
        return 0
