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
pdf_path = r"05-SUPPLEMENTARY RETIREMENT SCHEME-0171-May-24.pdf"
#pdf_path = r"06-SUPPLEMENTARY RETIREMENT SCHEME-0171-Jun-24.pdf"
#pdf_path = r"07-SUPPLEMENTARY RETIREMENT SCHEME-0171-Jul-25.pdf"
output_excel = r"srs_transaction_details.xlsx"

# 2. Helper: parse one line of transaction text
#    We only keep Date, Description and Credit amount

def parse_transaction_line(line: str):
    """Parse a single transaction line.

    Only keep rows that look like dividend credits, for example:
      '13MAY CR DIVIDENDS FOR 92FC 13.75 19,067.06'
      '13MAY CR DIVIDENDS FOR SGX 34.00 19,101.06'
      '21MAY CR DIVIDENDS FOR HAWP 100.00 17,024.11'

    Pattern expected:
        <DD><MON> CR DIVIDENDS FOR <something> <amount> <balance>
    """

    # Strict pattern:
    #   1) Date token at the start, like 13MAY
    #   2) Description that must contain 'CR DIVIDENDS FOR ...'
    #   3) Two monetary values at the end: amount then balance
    m = re.match(
        r"^(\d{2}[A-Z]{3})\s+(CR DIVIDENDS FOR\s+.+?)\s+([0-9,]+\.\d{2})\s+([0-9,]+\.\d{2})\s*$",
        line,
    )
   
    if not m:
        # Not a dividend credit row
        return None

    #print(f"Transaction lines extracted:{m}\n")

    date_raw, description, amount_str, _balance_str = m.groups()

    # Convert from '13MAY' to '13-May-2024'
    month_map = {
        "JAN": "Jan", "FEB": "Feb", "MAR": "Mar", "APR": "Apr",
        "MAY": "May", "JUN": "Jun", "JUL": "Jul", "AUG": "Aug",
        "SEP": "Sep", "OCT": "Oct", "NOV": "Nov", "DEC": "Dec",
    }
    day = date_raw[:2]
    mon = month_map.get(date_raw[2:5], date_raw[2:5])
    #IMPORTANT: The year is fixed to 2024. Change this if the year is different.
    date = f"{day}-{mon}-24"

    # For dividend rows, amount is always a credit
    return {
        "Date": date,
        "Description": description,
        "Credit ($)": amount_str,
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

with pdfplumber.open(pdf_path) as pdf:
    #print("pdf.pages")
    for page in pdf.pages:
        text = page.extract_text() or ""
        #print("Printing text")
        #print(f"text:{text}")
        
        if "TTRRAANNSSAACCTTIIOONN DDEETTAAIILLSS" not in text:
            continue

        # Keep only the part between "TRANSACTION DETAILS" and
        # the next section header "SECURITY INVESTMENT ACTIVITY"
        section = text.split("TTRRAANNSSAACCTTIIOONN DDEETTAAIILLSS", 1)[1]
        
        if "SECURITY INVESTMENT ACTIVITY" in section:
            section = section.split("SECURITY INVESTMENT ACTIVITY", 1)[0]

        for raw_line in section.splitlines():
            line = raw_line.strip()
            #print(f"Line:{line}\n")
            # Pass every line in the TRANSACTION DETAILS section to the parser.
            # The parser itself will decide if the line is a valid transaction row.
            record = parse_transaction_line(line)
            if record:
                rows.append(record)
                
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

print(f"Rows after cleaning:{rows}\n")

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
