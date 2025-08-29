# LITERATURE BIRGER NEUHAUS

# 1. Extract PDF name into an excel
import os
import pandas as pd

# chose the folder containing PDFs
pdf_folder = r"folder path" #change based on the folder path 

def get_pdfs(pdf_folder):
    # get PDF names without the ".pdf" extension
    pdfs = [os.path.splitext(f)[0] for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]
    return sorted(pdfs)

def save_pdfs_to_excel(pdf_folder, output_file_name="pdf_list_RZ.xlsx"): #modify the output file
    pdfs = get_pdfs(pdf_folder)
    # save excel in the same folder as PDFs
    output_file = os.path.join(pdf_folder, output_file_name)
    df = pd.DataFrame({"PDF Name": pdfs})
    df.to_excel(output_file, index=False)

# run the function
save_pdfs_to_excel(pdf_folder)




# 2. Put the full references list from word into an excel 
from docx import Document

# path to word file
docx_file = r"file path" #change based on Word the file path
doc = Document(docx_file)

references = []
current_ref = ""

for para in doc.paragraphs:
    text = para.text.strip()
    if not text:
        continue  # skip empty paragraphs
    if text.startswith("http"):
        # merge DOI/URL to previous reference (to avoid DOI placed in different cell)
        current_ref += " " + text
    else:
        if current_ref:
            references.append(current_ref)
        current_ref = text

# add the last reference
if current_ref:
    references.append(current_ref)

# save to Excel
df = pd.DataFrame({"Reference": references})
output_file = r"output folder path" # fill path to output file
df.to_excel(output_file, index=False)




# 3. Match the PDF filename and full reference (based on first author and year)
# matched pdf and reference move out to sheet 2
import re

file_path = r"excel file path" # path to excel in which pdf filename and reference together in the same sheet
df = pd.read_excel(file_path, header=None, dtype=str)

# ensure columns exist
while df.shape[1] < 3:
    df[df.shape[1]] = ""

# extract first author and year from reference
def extract_author_year(ref):
    try:
        first_author = ref.split(",")[0].strip()
        year_match = re.search(r"\((\d{4})\)", ref)
        year = year_match.group(1) if year_match else ""
        return first_author, year
    except:
        return "", ""

# list of all PDFs (in column C without column name)
pdfs = df[2].dropna().tolist()
used_pdfs = set()

# create new sheet for matched items
matched_rows = []

# loop through references
for idx, row in df.iterrows():
    ref = row[0] if pd.notna(row[0]) else ""
    first_author, year = extract_author_year(ref)
    matched_pdf = ""
    for pdf in pdfs:
        if pdf in used_pdfs:
            continue
        if first_author.lower() in pdf.lower() and year in pdf:
            matched_pdf = pdf
            used_pdfs.add(pdf)
            break
    if matched_pdf:
        matched_rows.append([ref, matched_pdf])
        # Remove from original sheet
        df.at[idx, 0] = ""
        df.at[idx, 2] = ""

# save updated Excel with two sheets
with pd.ExcelWriter(file_path.replace(".xlsx", "_matched_sheet2.xlsx")) as writer:
    df.to_excel(writer, sheet_name="Sheet1", index=False, header=False)
    pd.DataFrame(matched_rows).to_excel(writer, sheet_name="Sheet2", index=False, header=False)






