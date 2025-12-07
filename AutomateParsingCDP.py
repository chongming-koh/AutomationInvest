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
#pdf_path = r"Nov 2025.pdf"
#pdf_path = r"Sep 2025.pdf"
pdf_path = r"Aug 2025.pdf"
#pdf_path = r"Oct 2025.pdf"
#pdf_path = r"07-Jul 2025.pdf"
#pdf_path = r"06-Jun 2025.pdf"
#pdf_path = r"03-Mar 2025.pdf"

output_excel = r"CDP_transaction_details.xlsx"

# 2. Helper: parse one line of transaction text
#    We only keep Date, Description and amount

def parse_transaction_line(line: str):
    """
    Expected pattern (examples):
      '14/11/2025 SGX Interim Cash Dividend - 600 units @ SGD 0.1075 64.50'
      '24/11/2025 DBS Interim Cash Dividend - 442 units @ SGD 0.6 265.20'
      '28/11/2025 NETLINK NBN TR Interim Cash Dividend - 1,000 units @ SGD 0.0271 27.10'

    Capture:
      Date        -> 14/11/2025
      Description -> SGX Interim Cash Dividend - 600 units @ SGD 0.1075
      Amount      -> 64.50
    """

    # 1) Start with DD/MM/YYYY
    # 2) Description must contain "Dividend"
    # 3) Last token is Amount Paid (allow optional "-" sign)
    m = re.match(
    r"^(\d{2}/\d{2}/\d{4})\s+(.+(?:Dividend|Distribution).+?)\s+(-?[0-9,]+\.\d{2})\s*$",
    line,
    flags=re.IGNORECASE,
)


    if not m:
        # Not a dividend credit row
        return None
    '''
    #Debug Purposes
    if m:
        print("\n--- Regex Match Debug ---")
        for i, g in enumerate(m.groups(), start=1):
            print(f"Group {i}: {g}")
        print("-------------------------\n")
        '''

    date_raw, description, amount_str = m.groups()

    # Convert from '14/11/2025' to '14-Nov-25' (or whatever you prefer)
    day, month_num, year_full = date_raw.split("/")

    month_map = {
        "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
        "05": "May", "06": "Jun", "07": "Jul", "08": "Aug",
        "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec",
    }
    mon = month_map.get(month_num, month_num)
    year_short = year_full[-2:]   # "2025" -> "25"
    date = f"{day}-{mon}-{year_short}"

    return {
        "Date": date,
        "Description": description,
        "Credit ($)": amount_str,
    }


# 3. Open PDF and extract the TRANSACTION DETAILS block
CODE_TO_NAME = {
    "CAPITA CHINA TR": "CapitaLand Retail China Trust",
    "CAPLAND ASCOTT T": "ASCOTT RESIDENCE TRUST",
    "CAPLAND INDIA T": "Capitaland India Trust",
    "CAPLAND INTCOM T": "Capitaland Integrated Commercial Trust",
    "CL ASCENDAS REIT": "Capitaland Ascendas",
    "DBS": "DBS",
    "FIRST REIT": "First Reits",
    "F & N": "F&N",
    "FRASERS CPT TR": "Frasers Centrepoint",
    "HAW PAR": "Haw Par",
    "KEPPEL DC REIT": "Keppel DC",
    "MAPLETREE IND TR": "Mapletree Industrial Trust",
    "MPLTR PAN TR": "Mapletree Pan Asia",
    "NETLINK NBN TR": "NetLink NBN Trust",
    "OCBC BANK": "OCBC",
    "PARKWAYLIFE REIT": "Parkway Life REITS",
    "RAFFLES MEDICAL": "Raffles Medical",
    "SGX": "SGX",
    "STI ETF": "STI ETF",
    "THAIBEV": "Thai Beverage",
    "VICOM LTD": "Vicom",
    "HONGKONG LAND": "Hong Kong Land",
    "MANULIFEREIT USD": "Manulife US Reits",
}
# This is to update the description to the string that i prefers
def clean_description(raw_desc: str) -> str:
    """Strip fixed prefix and map the leading code to a full name.

    Examples:
      'PARKWAYLIFE REIT Interim Cash Dividend - 600 units @ SGD 0.0177'
        -> 'Parkway Life REITS'

      'CAPITA CHINA TR Final Cash Dividend - 6,800 units @ SGD 0.0264'
        -> 'CapitaLand Retail China Trust'
    """
    # Remove prefix and trim
    desc = raw_desc.replace("CR DIVIDENDS FOR", "").strip()

    upper_desc = desc.upper()

    # Match on leading code from CODE_TO_NAME
    for code, full_name in CODE_TO_NAME.items():
        if upper_desc.startswith(code.upper()):
            return full_name

    # If no match, keep the cleaned original text
    return desc

rows = []

with pdfplumber.open(pdf_path) as pdf:
    
    for page in pdf.pages:
        text = page.extract_text() or ""
        #print("Printing text")
        #print(f"text:{text}")
        
        if "Cash Transaction" not in text:
            continue

        # Keep only the part between "TRANSACTION DETAILS" and
        # the next section header "SECURITY INVESTMENT ACTIVITY"
        section = text.split("Cash Transaction", 1)[1]
        
        if "Your Securities Account is Linked To" in section:
            section = section.split("Your Securities Account is Linked To", 1)[0]

        for raw_line in section.splitlines():
            line = raw_line.strip()
            #print(f"Line:{line}\n")
            #print(f"{line}\n")

            # Pass every line in the TRANSACTION DETAILS section to the parser.
            # The parser itself will decide if the line is a valid transaction row.
            record = parse_transaction_line(line)
            if record:
                rows.append(record)
                
# Clean Description field before saving
# 1) Strip the fixed prefix 'CR DIVIDENDS FOR'
# 2) Map issuer code at the start of the description to full name via lookup table
for r in rows:
    original_desc = r["Description"]

    # Use helper to produce a clean security name
    cleaned_desc = clean_description(original_desc)
    r["Description"] = cleaned_desc

    # New fields for Excel
    r["Ticker"] = ""              # empty for now
    r["Year"] = "2025"            # fixed year
    r["Currency"] = "SGD"         # fixed currency
    r["Net Amount"] = r["Credit ($)"]  # same as Credit ($)

#print(f"Rows after cleaning:{rows}\n")

# 4. Save to Excel (Date, Description, Credit only)
# df = pd.DataFrame(rows, columns=["Date", "Description", "Credit ($)"])
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

[df_summary] = [
    df_summary
    .groupby(group_cols, as_index=False)
    .agg({
        "Credit ($)": "sum",
        "Net Amount": "sum",
    })
]

# Optional: round to 2 decimal places
df_summary["Credit ($)"] = df_summary["Credit ($)"].round(2)
df_summary["Net Amount"] = df_summary["Net Amount"].round(2)

# 6. Save to Excel with two sheets: Sheet1 (detailed), Sheet2 (summarized)
with pd.ExcelWriter(output_excel) as writer:
    df.to_excel(writer, sheet_name="Sheet1", index=False)
    df_summary.to_excel(writer, sheet_name="Sheet2", index=False)

print(f"Extracted {len(df)} detailed rows to {output_excel} (Sheet1)")
print(f"Summarized to {len(df_summary)} rows in Sheet2")
