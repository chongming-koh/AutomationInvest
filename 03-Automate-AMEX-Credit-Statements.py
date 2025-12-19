'''

This script is used to parse the dividends text from the PDF file and save the results to an Excel file.
uv init to create pyproject.toml, uv.lock and .venv/
Install packages with uv and not pip. uv add pdfplumber pandas
To run the scripy, uv run python ParseDividendsText.py
'''


import pdfplumber
import pandas as pd
import re
from pathlib import Path
from datetime import datetime

# 1) Paths
base_dir = Path(r"Amex")
pdf_files = sorted(base_dir.rglob("*.pdf"))
output_excel = r"AmexCreditCardDetails.xlsx"

# 2) Markers (robust for old/new PDFs)
start_marker_re = re.compile(r"Details\s+Foreign\s+Spending\s+Amount\s*S\$", re.IGNORECASE)
end_marker = "Total of New Transactions"

# 3) Regex
DATE_DDMMYY_DOTS_RE = re.compile(r"^\d{2}\.\d{2}\.\d{2}\b")

# Matches:
# 31.01.21 PAYMENT BY TELEPHONE/INTERNET BANKING 723.40
# 19.01.21 XTRA AMK HUB SINGAPORE 25.35
# and also handles amount with trailing CR on same line: 723.40CR
TX_RE = re.compile(
    r"^(?P<date>\d{2}\.\d{2}\.\d{2})\s+"
    r"(?P<desc>.+?)\s+"
    r"(?P<amt>[0-9,]+\.\d{2})(?P<cr>CR)?\s*$"
)

def format_date_and_year(date_str: str) -> tuple[str, str]:
    """
    Input:  dd.mm.yy  (example: 25.10.20)
    Output: ("25 OCT", "2020")
    """
    dt = datetime.strptime(date_str, "%d.%m.%y")
    day_mon = dt.strftime("%d %b").upper()   # "04 NOV"
    year_yyyy = dt.strftime("%Y")            # "2020"
    return day_mon, year_yyyy

def format_amount_for_excel(amount_str: str, is_credit: bool) -> str:
    """
    amount_str: like "36.12" (may contain commas)
    is_credit: True if CR applies
    Returns: "(36.12)" for credit, else "36.12"
    """
    amt = (amount_str or "").replace(",", "").strip()
    return f"({amt})" if is_credit else amt

def is_redundant_line(line: str) -> bool:
    s = (line or "").strip()
    if not s:
        return True

    if s == end_marker or start_marker_re.search(s):
        return True

    # Common AMEX statement noise (add more phrases here if you encounter them)
    redundant_substrings = [
        "American Express International",
        "UEN",
        "Statement of Account",
        "Prepared for Membership Number",
        "Membership Number",
        "PAYMENT ADVICE",
        "Please return",
        "Minimum Payment",
        "Due by",
        "Enter amount enclosed",
        "Please make crossed cheque payable",
        "AMERICAN EXPRESS",
        "Please do not write",
        "The Rewards Card",
        "Page ",
        "Important Information",
        "Foreign Currency Charges",
        "Online Services",
        "Payment Method",
        "Privacy:",
        "Limited Liability",
        "Credit Card Interest Rate Policy",
        "log on to americanexpress.com.sg",
        "amex.co/",
        "reply envelope",
    ]

    s_lower = s.lower()
    return any(sub.lower() in s_lower for sub in redundant_substrings)

def parse_transaction_line(line: str):
    """
    Returns dict with:
      - Transaction Date Captured  (dd MMM, example: 04 NOV)
      - Year                       (yyyy, example: 2020)
      - Description Captured
      - Amount Captured            ((36.12) for CR, else 36.12)
    Or None if the line is not a transaction row.
    """
    m = TX_RE.match(line.strip())
    if not m:
        return None

    raw_date = m.group("date")
    try:
        tx_date, year_yyyy = format_date_and_year(raw_date)
    except Exception:
        tx_date, year_yyyy = raw_date, ""

    amt = format_amount_for_excel(m.group("amt"), bool(m.group("cr")))

    return {
        "Transaction Date Captured": tx_date,
        "Year": year_yyyy,
        "Description Captured": m.group("desc").strip(),
        "Amount Captured": amt,
    }

def extract_marker_segments(all_text: str) -> list[str]:
    # normalize common invisible spaces
    all_text = all_text.replace("\xa0", " ")

    em = re.escape(end_marker)

    # Find: start_marker ... until next start_marker or end_marker
    seg_re = re.compile(
        rf"(?:{start_marker_re.pattern})(.*?)(?=(?:{start_marker_re.pattern})|{em})",
        flags=re.DOTALL | re.IGNORECASE
    )
    return [m.group(1) for m in seg_re.finditer(all_text)]

def extract_rows_from_pdf(pdf_path: Path) -> list[dict]:
    """Extract transaction rows from one PDF. Returns a list of dict rows."""
    with pdfplumber.open(str(pdf_path)) as pdf:
        all_text = "\n".join((page.extract_text() or "") for page in pdf.pages)

    # Pull all segments that start after the marker, across repeated pages
    segments = extract_marker_segments(all_text)
    if not segments:
        print(f"[SKIP] Markers not found in: {pdf_path}")
        return []

    # Build clean lines from all segments
    raw_lines: list[str] = []
    for seg in segments:
        raw_lines.extend([ln.strip() for ln in seg.splitlines()])

    clean_lines: list[str] = []
    for ln in raw_lines:
        s = (ln or "").strip()
        if not s:
            continue
        if is_redundant_line(s):
            continue
        clean_lines.append(s)

    rows: list[dict] = []
    i = 0
    while i < len(clean_lines):
        line = clean_lines[i].strip()

        # Standalone "CR" line belongs to the previous transaction amount
        if line == "CR":
            if rows and isinstance(rows[-1].get("Amount Captured"), str):
                prev = rows[-1]["Amount Captured"].strip()
                if not (prev.startswith("(") and prev.endswith(")")):
                    rows[-1]["Amount Captured"] = format_amount_for_excel(prev, True)
            i += 1
            continue

        # Only send plausible transaction lines to the parser
        if not DATE_DDMMYY_DOTS_RE.match(line):
            i += 1
            continue

        rec = parse_transaction_line(line)
        if not rec:
            i += 1
            continue

        # If next line is "CR", attach it to amount (convert to parentheses)
        if i + 1 < len(clean_lines) and clean_lines[i + 1].strip() == "CR":
            rec["Amount Captured"] = format_amount_for_excel(rec["Amount Captured"], True)
            i += 1

        rows.append(rec)
        i += 1

    print(f"[OK] {pdf_path} -> {len(rows)} rows")
    return rows

# 4) Process all PDFs
all_rows: list[dict] = []
for pdf_file in pdf_files:
    all_rows.extend(extract_rows_from_pdf(pdf_file))

# 5) Save to Excel (single sheet, 4 columns with Year in between)
df = pd.DataFrame(
    all_rows,
    columns=["Transaction Date Captured", "Year", "Description Captured", "Amount Captured"]
)
df.to_excel(output_excel, sheet_name="Transactions", index=False)

print(f"Total PDFs scanned: {len(pdf_files)}")
print(f"Total rows extracted: {len(df)}")
print(f"Saved combined output to: {output_excel}")


