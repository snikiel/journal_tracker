# citation_parser
A tool for analyzing and formatting journal citations

## Getting Started
### Dependencies
1. `git` - source control system for retrieving the source code to run it.
1. `python` - executable to run the scripts which are in python.
1. `uv` - the package and environment manager for python.
    1. `python3 -m ensurepip --upgrade`
    1. `pip install uv --upgrade`
### To run this script, you must:
1. Clone this repository from github `git clone https://github.com/snikiel/journal_tracker.git`
1. Copy your xlsx data file into the project's `/data/` folder.
1. Run `uv run python main.py --input [the_name_of_your_data_file_from_step_2.xlsx]` from the project's root folder, this will prompt and/or handle the installation of a handfull of modules.
1. You can also include the `--output` flag followed by whatever file name you would like your parsed data file to be saved as. This will be saved alongside the input file in the project's `/data/` folder.
1. This will first automatically run the suite of test cases, and if they all pass, it will proceed to parse and validate the data and log any anomolies, producing a parsed file.
1. Each time you run the script, a new log file will be created in the project's `/log/` folder, so you can review it after the fact.

## Specifications
### Overview
Citations from various journals need to be cut apart so that various parts of them can be sorted, filtered, and analyzed.

### Input
An .xlsx file that contains multiple sheets, each with a single column that contains:
1. Author(s): Last Name, Initials, followed by a comma and an ampersand before the last author in the list
1. Year of publish
1. Article title
1. Journal name
1. Volume of the journal
1. Issue of that volume
1. Pages within that issue
1. A URL to the article on doi.org

There may be additional columns to be passed through. #TBD

### Output
An .xlsx file with the same sheets with the following columns:
1. The original value passed in from the input file
1. Each of the defined input data split out into their own separate columns
1. A calculated number of how many authors were listed for the paper
1. Word count of the article? #TBD
1. Linkages between articles? #TBD

### Requirements
1. The code should be passed along so that it can be at least reviewed, and possibly re-used.
1. To that end, this should be put up on github.
1. This should also use some kind of virtual environment to help document and handle the versioning and dependency list.
1. Tests should be included to verify that each of the edge cases is handled correctly.
1. Putting as much code into modules as possible will help with this.