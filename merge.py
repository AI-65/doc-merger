import os
import subprocess
from PyPDF2 import PdfMerger, PdfReader

# Set the path to the directory containing the DOC files
doc_dir = r'/path/to/doc/folder'

# Set the path to the LibreOffice 'soffice' executable
soffice_path = r'C:\Program Files\LibreOffice\program\soffice.exe'  # Modify this path as needed

# Create a list of all the DOC files in the directory
doc_files = [os.path.join(doc_dir, f) for f in os.listdir(doc_dir) if f.endswith('.doc')]

# Create a new PdfMerger object
merger = PdfMerger()

# Loop through the DOC files and convert them to PDFs, then add them to the merger
for doc_file in doc_files:
    pdf_file = doc_file.replace('.doc', '.pdf')
    subprocess.run([soffice_path, '--headless', '--convert-to', 'pdf', doc_file, '--outdir', doc_dir])
    with open(pdf_file, 'rb') as f:
        merger.append(PdfReader(f))

# Merge the PDFs and save the output to a file
output_path = os.path.join(doc_dir, 'merged.pdf')
with open(output_path, 'wb') as f:
    merger.write(f)

# Clean up the temporary PDF files
for doc_file in doc_files:
    pdf_file = doc_file.replace('.doc', '.pdf')
    os.remove(pdf_file)
