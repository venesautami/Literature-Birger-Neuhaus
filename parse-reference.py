# 4. Parsing the full reference 
# the logic need to be adjusted based on the full reference format
import pandas as pd
import re

file_path = r"excel path file"
df = pd.read_excel(file_path)

def parse_reference(ref):
    # extract DOI first
    doi_match = re.search(r"(https?://\S+)", ref)
    doi = doi_match.group(1) if doi_match else ""
    # remove everything after DOI (except DOI itself)
    if doi:
        ref = ref.split(doi)[0] + ' ' + doi

    # authors
    authors_match = re.match(r"^(.*?)\s\(\d{4}\)", ref)
    authors = authors_match.group(1).strip() if authors_match else ""
    authors = re.sub(r"\s*&\s*", "; ", authors)

    # year
    year_match = re.search(r"\((\d{4})\)", ref)
    year = year_match.group(1) if year_match else ""

    # remaining after year
    title_start = ref.find(f"({year})") + len(f"({year})") if year else 0
    remaining_text = ref[title_start:].strip()

    # split title and journal/suffix
    title = ""
    journal = ""
    suffix = ""
    match = re.match(r"(.*?\.)\s*([^.]+?),\s*(\d+),\s*([\d–\-]+)\.?", remaining_text)
    if match:
        title = match.group(1).strip()
        journal = match.group(2).strip()
        vol = match.group(3).strip()
        pages = match.group(4).strip()
        suffix = f"{vol}: {pages}"
    else:
        parts = remaining_text.rsplit('.', 1)
        title = parts[0].strip() if parts else remaining_text.strip()

    return pd.Series([title, authors, year, journal, suffix, doi])

# apply parsing
df[['title', 'authors', 'year', 'journal', 'suffix', 'doi']] = df['references'].apply(parse_reference) #fill with column name

# save back to the same Excel file
df.to_excel(file_path, index=False)


# Save back to the same Excel file, not a new excel be careful with losing the original version
df.to_excel(file_path, index=False)
print("References parsed and Excel updated successfully!")

