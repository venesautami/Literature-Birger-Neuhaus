# Literature PDF & Reference Manager
To automate the management of literature references and PDFs, the initial purpose was for bulk uploading a collection of literature with full reference lists and available PDFs. The full references then need to be parsed.

## Steps
1. Extract PDF file names from a folder into an Excel sheet.
2. Extract references from a Word document into an Excel sheet.
3. Match PDF filenames with references based on first author and year.
4. Parse full references into structured Excel columns, including:
   - Authors
   - Year
   - Title
   - Journal
   - Volume/Pages (suffix)
   - DOI

## Script
pdf-reference-match.py (for steps 1-3)
parse-reference.py (for step 4)

## Development Note
Some parts of the scripts and parsing logic were improved with the assistance of AI tools to optimize code structure and automation.
