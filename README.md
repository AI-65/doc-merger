# DOC Merger

A Python script to merge an arbitrarily large set of Microsoft Word (.doc) files into a single PDF file using PyPDF2 and LibreOffice.

## Dependencies

- [Python](https://www.python.org/downloads/)
- [PyPDF2](https://pypi.org/project/PyPDF2/)
- [LibreOffice](https://www.libreoffice.org/download/download/)

## Installation

1. Install [Python](https://www.python.org/downloads/) on your system.
2. Install the required Python package using pip: pip install PyPDF2
3. Install [LibreOffice](https://www.libreoffice.org/download/download/) and make sure `soffice` is available in your system's PATH.

## Usage

1. Modify the `merge.py` script to set the `doc_dir` variable to the directory containing your `.doc` files.
2. Set the `soffice_path` variable to the path of the `soffice.exe` file in your LibreOffice installation.
3. Run the script: python merge.py


The script will convert all `.doc` files in the specified directory to `.pdf` files, merge them into a single PDF named `merged.pdf`, and save the output in the same directory.




