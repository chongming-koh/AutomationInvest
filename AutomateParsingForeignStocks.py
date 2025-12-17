'''

This script is used to parse the dividends text from the PDF file and save the results to an Excel file.
uv init to create pyproject.toml, uv.lock and .venv/
Install packages with uv and not pip. uv add pdfplumber pandas
To run the scripy, uv run python ParseDividendsText.py
'''


import pdfplumber
import pandas as pd
import re

# 1. Set your paths
#pdf_path = r"1Page.pdf"
pdf_path = r"Statement-28112025M2201962-s.pdf"
#pdf_path = r"06-SUPPLEMENTARY RETIREMENT SCHEME-0171-Jun-24.pdf"
#pdf_path = r"07-SUPPLEMENTARY RETIREMENT SCHEME-0171-Jul-25.pdf"
output_excel = r"foreign_transaction_details.xlsx"

# 2. Helper: parse one line of transaction text
#    We only keep Date, Description and Credit amount

def parse_transaction_line(line: str):
    """
    Parse a single transaction line for:

      - 'CR Note W.E.F ...'
      - 'DR Note HANDLING ...'

    Examples:
      '19/11/2025 CRC7789419 CR Note W.E.F 22 MAY 31.07 31.07'
      '19/11/2025 DRC7789419 DR Note HANDLING 4.20 26.87'
      '03/11/2025 CRC7780886 CR Note W.E.F 22 MAY 1,000.00 1,000.00'
      '03/11/2025 DRC7780886 DR Note HANDLING CHARGES-HEIM 17.71 982.29'
      '11/11/2025 CRC7784882 CR Note W.E.F 22 MAY 48.64 48.64'
      '11/11/2025 DRC7784882 DR Note HANDLING 3.18 45.46'

    Any other 'Note ...' lines (for example 'Note INTEREST CREDIT BIL')
    are ignored.
    """

    # Hard filter so we never pick up 'Note INTEREST CREDIT BIL' etc.
    if "CR Note W.E.F" not in line and "DR Note HANDLING" not in line:
        return None

    m = re.match(
        r"^"
        r"(\d{2}/\d{2}/\d{4})"      # 1: date in DD/MM/YYYY
        r"\s+([A-Z0-9]+)"           # 2: reference like CRC7789419 / DRC7780886
        r"\s+(CR|DR)"               # 3: CR/DR indicator
        r"\s+Note\s+(.+?)"          # 4: free text after 'Note ' up to the amounts
        r"\s+([0-9,]+\.\d{2})"      # 5: amount
        r"\s+([0-9,]+\.\d{2})"      # 6: balance
        r"\s*$",
        line,
    )

    if not m:
        # Line did not match the strict pattern
        return None

    date_str, ref_no, crdr_flag, note_suffix, amount_str, balance_str = m.groups()

    # Build a readable description, but you can tweak this
    description = f"{crdr_flag} Note {note_suffix}"

    # Keep the keys your downstream code expects
    return {
        "Date": date_str,
        "Description": description,
        "Credit ($)": amount_str,
        # Extra fields if you want to use them later
        "Ref": ref_no,
        "Balance": balance_str,
        "CRDR": crdr_flag,
    }




# 3. Open PDF and extract the TRANSACTION DETAILS block
CODE_TO_NAME = {
    "1EG0": "Nikko AM SGD Corp Bond ETF",
    "SBIETF": "ABF SINGAPORE Bond Index",
    "STETF": "STI ETF",
    "1DH9": "NetLink NBN Trust",
    "4H4M": "ASCOTT RESIDENCE TRUST",
    "92FC": "Vicom",
    "ASCEND": "Capitaland India Trust",
    "CMT": "Capitaland Integrated Commercial Trust",
    "CRCT": "CapitaLand Retail China Trust",
    "F-LITR": "Frasers Logistics Commerical",
    "HAWP": "Haw Par",
    "KDCRET": "Keppel DC",
    "MCT": "Mapletree Pan Asia",
    "MITR": "Mapletree Industrial Trust",
    "OCBCBK": "OCBC",
    "PKREIT": "Parkway Life REITS",
    "RMEDIC": "Raffles Medical",
    "SGX": "SGX",
    "STEL": "Singtel",
}
rows = []

start_marker = "S t a t e m e n t O f A c c o u n t"
end_marker   = "C u s t o d y S t a t e m e n t"

with pdfplumber.open(pdf_path) as pdf:
    # 1) Concatenate text from all pages
    all_text = ""
    for page in pdf.pages:
        page_text = page.extract_text() or ""
        all_text += page_text + "\n"

#print(all_text)

# 2) Cut from "Statement Of Account" to "Custody Statement"
if start_marker not in all_text:
    raise ValueError(f"Start marker '{start_marker}' not found in PDF text")

section = all_text.split(start_marker, 1)[1]

if end_marker in section:
    section = section.split(end_marker, 1)[0]
else:
    # Optional: warn if end marker missing instead of raising
    print(f"Warning: end marker '{end_marker}' not found, using rest of document")

#print(section)

# 3) Now parse only the lines inside this section,
#    and stitch short non-date lines onto the previous description.
date_re = re.compile(r"^\d{2}/\d{2}/\d{4}\b")

lines = [l.strip() for l in section.splitlines()]
i = 0
while i < len(lines):
    line = lines[i]

    # Only attempt parsing for lines that start with a date
    if date_re.match(line):
        record = parse_transaction_line(line)
        if record:
            # Look ahead to see if the next line is a short continuation
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()

                # Continuation rule:
                # - next line does not start with a date
                # - next line is not empty
                # - next line is relatively short (e.g. ≤ 40 chars)
                if (
                    next_line
                    and not date_re.match(next_line)
                    and len(next_line) <= 40
                ):
                    # Append the continuation text to the description
                    record["Description"] += " " + next_line
                    # Skip the continuation line in the main loop
                    i += 1

            rows.append(record)

    # Move to the next line
    i += 1


                
#print(f"Rows before cleaning:{rows}\n")

# Clean Description field before saving
# 1) Strip the fixed prefix 'CR DIVIDENDS FOR'
# 2) Map short codes (92FC, HAWP, etc.) to full names via lookup table
for r in rows:
    desc = r["Description"]

    # Remove prefix and trim
    desc = desc.replace("CR DIVIDENDS FOR", "").strip()

    # Map short code to full name if available
    full_name = CODE_TO_NAME.get(desc, desc)
    r["Description"] = full_name

    # New fields for Excel
    r["Ticker"] = ""              # empty for now
    r["Year"] = "2024"            # fixed year
    r["Currency"] = "SGD"         # fixed currency
    r["Net Amount"] = r["Credit ($)"]  # same as Credit ($)

#print(f"Rows after cleaning:{rows}\n")

# 4. Save to Excel (Date, Description, Credit only)
#df = pd.DataFrame(rows, columns=["Date", "Description", "Credit ($)"])
df = pd.DataFrame(rows)[["Description", "Ticker", "Year", "Date", "Currency", "Credit ($)", "Net Amount"]]


# 5. Build summarized DataFrame (Sheet2)
# Convert Credit ($) and Net Amount to numeric so they can be summed
df_summary = df.copy()

for col in ["Credit ($)", "Net Amount"]:
    # Remove commas if any, then convert to float
    df_summary[col] = (
        df_summary[col]
        .astype(str)
        .str.replace(",", "", regex=False)
        .astype(float)
    )

# Group by Description and other identifying fields
# This collapses multiple payouts for the same company on the same date
group_cols = ["Description", "Ticker", "Year", "Date", "Currency"]

df_summary = (
    df_summary
    .groupby(group_cols, as_index=False)
    .agg({
        "Credit ($)": "sum",
        "Net Amount": "sum",
    })
)

# Optional: round to 2 decimal places
df_summary["Credit ($)"] = df_summary["Credit ($)"].round(2)
df_summary["Net Amount"] = df_summary["Net Amount"].round(2)

# 6. Save to Excel with two sheets: Sheet1 (detailed), Sheet2 (summarized)
with pd.ExcelWriter(output_excel) as writer:
    df.to_excel(writer, sheet_name="Sheet1", index=False)
    df_summary.to_excel(writer, sheet_name="Sheet2", index=False)

print(f"Extracted {len(df)} detailed rows to {output_excel} (Sheet1)")
print(f"Summarized to {len(df_summary)} rows in Sheet2")
