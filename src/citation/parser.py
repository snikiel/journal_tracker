"""
Takes an Excel file of journal citations and parses them 
into a new Excel file with structured fields.
"""
import re
import pandas as pd

class CitationParser():
    """
    Class to parse journal citations from an Excel file and output structured data.
    """
    def __init__(self, input_file, output_file):
        self.input_file = input_file
        self.output_file = output_file

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
        xl = pd.ExcelFile('Journal Tracking.xlsx')

        # Open new file for writing, so we can loop through all sheets
        # without knowing/caring what they're called.
        with pd.ExcelWriter('parsed_full.xlsx') as writer:
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


                    entries.append({
                        'Original Citation': line,
                        'Authors': authors,
                        'Year': year,
                        'Title': title,
                        'Journal': journal,
                        'VolumeIssue': volissue,
                        'Pages': pages,
                        'DOI': doi,
                        'Number of Authors': self.num_authors(authors),
                    })

                # Convert to DataFrame and export
                df = pd.DataFrame(entries)
                df.to_excel(writer, sheet_name=str(sheet_name), index=False, engine='openpyxl')

                print(f'created sheet {sheet_name}')
