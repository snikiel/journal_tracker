"""

"""
from pprint import pformat
from numpy import floor

class Citation():
    """
    Class to represent a single citation with structured fields.
    """
    def __init__(self, original_citation, logger, row=None, sheet_name=None, authors=None,
                 year=None, title=None, journal=None, volume_issue=None, pages=None, doi=None,
                 author_num=None):
        self.original_citation = original_citation
        self.row = row
        self.sheet_name = sheet_name
        self.authors = authors
        self.year = year
        self.title = title
        self.journal = journal
        self.volume_issue = volume_issue
        self.pages = pages
        self.doi = doi
        self.author_num= author_num
        self.parse_errors = False
        self.logger = logger

    def __str__(self):
        return pformat(vars(self))

    def to_dict(self) -> dict:
        """
        Convert the Citation object to a dictionary.
        """
        attrs = vars(self)
        del attrs['logger']
        del attrs['row']
        return attrs

    def num_authors(self):
        """Takes a string of authors and returns the number of authors"""
        if not self.authors:
            print('No authors found in citation: "%s", setting number of authors to 0',
                  self.original_citation)
            self.author_num = 0
            return
        omit_suffixes = ["Jr", "Sr", "Jr.", "Sr.", "III", "IV"]
        names = self.authors.split(',')
        filtered_names = [name for name in names if name.strip() not in omit_suffixes]
        res = len(filtered_names)/2

        if not res.is_integer():
            print(f'"{self.authors}" counted {res} authors')
        self.author_num = res

    def validate(self):
        """
        Validate the parsed citation fields and print warnings for any issues.
        """
        if not self.year:
            self.log_parsing_warning("No year found")
        if not self.title:
            self.log_parsing_warning("No title found")
        if not self.journal:
            self.log_parsing_warning("No journal found")
        if self.journal and len(self.journal) > 0 and self.journal[0] != self.journal[0].upper():
            self.log_parsing_warning("Journal name doesn't match sheet name")
        if not self.volume_issue:
            self.log_parsing_warning("No volume/issue found")
        if not self.pages:
            self.log_parsing_warning("No pages found")
        if not self.doi:
            self.log_parsing_warning("No DOI found")
        if self.author_num is not None and self.author_num != floor(self.author_num):
            self.log_parsing_warning("Number of authors is not an integer")

        if self.parse_errors:
            self.logger.debug(
                '   Debug:   (sheet: "%s", row: %s) - Citation: %s', 
                self.sheet_name, self.row, self)

    def log_parsing_warning(self, message):
        """
        Log a parsing warning with details about the citation and its location.
        """
        self.parse_errors = True
        self.logger.warning(
            'Warning: (sheet: "%s", row: %s) - %s in citation.', 
            self.sheet_name, self.row, message)
