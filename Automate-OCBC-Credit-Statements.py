'''

This script is used to parse the dividends text from the PDF file and save the results to an Excel file.
uv init to create pyproject.toml, uv.lock and .venv/
Install packages with uv and not pip. uv add pdfplumber pandas
To run the scripy, uv run python ParseDividendsText.py
'''


import pdfplumber
import pandas as pd
import re


# 1) Paths
from pathlib import Path
base_dir = Path(r"OCBC")
pdf_files = sorted(base_dir.rglob("*.pdf"))
output_excel = r"OCBCCreditCardDetails.xlsx"

# 2) Markers
start_marker = "TRANSACTION DATE DESCRIPTION AMOUNT (SGD)"
end_marker = "SUBTOTAL"

# NEW: Month formatting for "DD/MM" -> "DD Mmm" (e.g., "02/04" -> "02 Apr")
MONTH_ABBR = {
    "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
    "05": "May", "06": "Jun", "07": "Jul", "08": "Aug",
    "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec",
}

def format_tx_date(date_str: str) -> str:
    # Accepts 'DD/MM' (or 'DD/MM/YYYY') and returns 'DD Mmm'
    s = (date_str or "").strip()
    parts = s.split("/")
    if len(parts) < 2:
        return s
    dd = parts[0].zfill(2)
    mm = parts[1].zfill(2)
    return f"{dd} {MONTH_ABBR.get(mm, mm)}"

# 3) Helpers
DATE_START_RE = re.compile(r"^\d{2}/\d{2}\b")

def is_redundant_line(line: str) -> bool:
    """Return True for statement headers/footers and other noise."""
    s = (line or "").strip()
    if not s:
        return True

    # Drop markers and repeated headers
    if s in {start_marker, end_marker}:
        return True

    redundant_substrings = [
        "OCBC 365 CREDIT CARD",
        "penalty interest rate",
        "back of statement",
        "TAN ",
        "xxxx-xxxx-xxxx-xxxx",
        "LAST MONTH'S BALANCE",
        "Co.Reg.no.:",
        "PAGE ",
        "OCBC Bank - Credit Cards",
        "65 Chulia Street",
        "OCBC Centre",
        "Singapore 049513",
        "CONTACT US",
        "1800 363 3333",
        "(65) 6363 3333",
        "when overseas",
    ]

    return any(sub in s for sub in redundant_substrings)


# Matches lines like:
# 14/03 -4751 BUS/MRT ... 3.95
# 02/04 PAYMENT BY INTERNET (241.75)
# 28/03 CASH REBATE (0.65)
TX_RE = re.compile(
    r"^(?P<date>\d{2}/\d{2})\s+"               # DD/MM
    r"(?P<desc>.+?)\s+"                        # description
    r"(?P<amt>\(?-?[0-9,]+\.\d{2}\)?)\s*$"    # 0.00 or (0.00) or -0.00
)


def parse_transaction_line(line: str):
    m = TX_RE.match(line.strip())
    if not m:
        return None

    return {
        "TRANSACTION DATE": format_tx_date(m.group("date")),   # UPDATED
        "DESCRIPTION": m.group("desc").strip(),
        "AMOUNT (SGD)": m.group("amt"),
    }

def get_subfolder_name(pdf_path: Path) -> str:
    """
    Returns the first-level subfolder under base_dir that contains this pdf.
    Example: OCBC/2023/xxx.pdf -> "2023"
    If the pdf is directly under OCBC/, it returns "OCBC".
    """
    rel_parts = pdf_path.relative_to(base_dir).parts
    if len(rel_parts) >= 2:
        return rel_parts[0]
    return base_dir.name

def extract_rows_from_pdf(pdf_path: Path) -> list[dict]:
    """Extract transaction rows from one PDF. Returns a list of dict rows."""

    year_value = get_subfolder_name(pdf_path)

    with pdfplumber.open(str(pdf_path)) as pdf:
        all_text = "\n".join((page.extract_text() or "") for page in pdf.pages)

    # Extract the block from header to SUBTOTAL (inclusive)
    block_re = re.compile(
        rf"{re.escape(start_marker)}.*?{re.escape(end_marker)}",
        flags=re.DOTALL
    )

    m = block_re.search(all_text)
    if not m:
        # No block found in this PDF, skip it
        print(f"[SKIP] Markers not found in: {pdf_path}")
        return []

    section = m.group(0)

    # Clean lines and ignore page-break repeated headers
    raw_lines = [ln.strip() for ln in section.splitlines()]

    clean_lines = []
    for ln in raw_lines:
        s = ln.strip()

        if s == start_marker:
            continue
        if s == end_marker:
            break
        if is_redundant_line(s):
            continue

        clean_lines.append(s)

    # Parse transactions, stitching wrapped descriptions when needed
    rows = []
    i = 0
    while i < len(clean_lines):
        line = clean_lines[i]

        rec = parse_transaction_line(line)
        if rec:
            # Add Year (subfolder name) between date and description
            rec["Year"] = year_value
            
            # Continuation: next line is not another transaction line, append to description
            if i + 1 < len(clean_lines):
                nxt = clean_lines[i + 1]
                if not TX_RE.match(nxt) and not is_redundant_line(nxt):
                    rec["DESCRIPTION"] = (rec["DESCRIPTION"] + " " + nxt).strip()
                    i += 1

            rows.append(rec)

        i += 1

    print(f"[OK] {pdf_path} -> {len(rows)} rows")
    return rows
    
# 4) Read PDF text
all_rows = []
pdf_files = sorted(base_dir.rglob("*.pdf"))

for pdf_file in pdf_files:
    all_rows.extend(extract_rows_from_pdf(pdf_file))

# 5) Save to Excel (single sheet, 3 columns)
df = pd.DataFrame(all_rows, columns=["TRANSACTION DATE", "Year", "DESCRIPTION", "AMOUNT (SGD)"])
df.to_excel(output_excel, sheet_name="Transactions", index=False)

print(f"Total PDFs scanned: {len(pdf_files)}")
print(f"Total rows extracted: {len(df)}")
print(f"Saved combined output to: {output_excel}")

