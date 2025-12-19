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

# 1) Paths
base_dir = Path(r"UOB")
pdf_files = sorted(base_dir.rglob("*.pdf"))
output_excel = r"UOBCreditCardDetails.xlsx"

# 2) Markers (your requested markers)
start_marker = "Post Trans Description of Transaction Transaction Amount"
end_marker = "SUB TOTAL"

# 3) Regex helpers

# Flexible start marker matcher to survive line breaks / spacing differences
# Replace START_MARKER_RE + END_MARKER_RE with these
START_MARKER_RE = re.compile(
    r"Post\s*(?:Date\s*)?Trans\s*(?:Date\s*)?Description\s*of\s*Transaction\s*Transaction\s*Amount",
    flags=re.IGNORECASE,
)
END_MARKER_RE = re.compile(r"SUB\s*TOTAL", flags=re.IGNORECASE)

def _compact(s: str) -> str:
    """Remove all whitespace to survive PDFs that glue words together."""
    return re.sub(r"\s+", "", s or "")

def extract_section_text(pdf_path: Path) -> str | None:
    """Return text between transaction table header and SUB TOTAL across all pages."""
    with pdfplumber.open(str(pdf_path)) as pdf:
        all_text = "\n".join((page.extract_text() or "") for page in pdf.pages)

    # 1) Try regex on raw text
    m_start = START_MARKER_RE.search(all_text)

    # 2) If not found, try matching on "compacted" text (no whitespace)
    if not m_start:
        compact_text = _compact(all_text)
        compact_start = _compact(start_marker)

        idx = compact_text.find(compact_start)
        if idx == -1:
            # also try a compacted regex pattern (same idea as START_MARKER_RE)
            # simplest fallback: look for "PostTransDescriptionofTransactionTransactionAmount"
            fallback = "PostTransDescriptionofTransactionTransactionAmount".lower()
            idx = compact_text.lower().find(_compact(fallback))

        if idx == -1:
            print(f"[SKIP] Start marker not found in: {pdf_path}")
            return None

        # Map compact index back to raw index by walking characters
        # Build raw_index_from_compact_index once
        raw_index_from_compact = []
        for raw_i, ch in enumerate(all_text):
            if not ch.isspace():
                raw_index_from_compact.append(raw_i)

        start_raw_idx = raw_index_from_compact[idx]
        after_start = all_text[start_raw_idx:]
    else:
        after_start = all_text[m_start.start():]

    # Find end marker after the start
    m_end = END_MARKER_RE.search(after_start)
    if not m_end:
        print(f"[SKIP] End marker not found in: {pdf_path}")
        return None

    return after_start[:m_end.end()]


# Transaction start line:
# "07 JUN 07 JUN CR INTEREST 16.84 CR"
TX_START_RE = re.compile(r"^(?P<post>\d{2}\s+[A-Z]{3})\s+(?P<trans>\d{2}\s+[A-Z]{3})\s+(?P<rest>.+)$")

# Amount token at end of a line, with optional CR/DR (with or without space)
AMT_AT_END_RE = re.compile(r"(?P<num>\d{1,3}(?:,\d{3})*\.\d{2})\s*(?P<suffix>CR|DR)?$", flags=re.IGNORECASE)

# Amount-only line (e.g., "7.00")
AMT_ONLY_RE = re.compile(r"^(?P<num>\d{1,3}(?:,\d{3})*\.\d{2})\s*(?P<suffix>CR|DR)?$", flags=re.IGNORECASE)

def normalize_amount(num: str, suffix: str | None) -> str:
    """
    - If suffix is CR, output (num)
    - If no suffix, output num
    - (Keeps DR as-is unless you want a different rule)
    """
    num = (num or "").replace(",", "").strip()
    suf = (suffix or "").strip().upper()

    if suf == "CR":
        return f"({num})"
    return num if not suf else f"{num}{suf}"


def is_redundant_line(line: str) -> bool:
    s = (line or "").strip()
    if not s:
        return True

    s_upper = s.upper()
    s_compact = re.sub(r"\s+", "", s).upper()

    # Always drop explicit markers / repeated headers
    if START_MARKER_RE.search(s) or END_MARKER_RE.search(s):
        return True
    if s in {start_marker, end_marker, "Date Date SGD"}:
        return True

    # Drop previous balance row
    if s_upper.startswith("PREVIOUS BALANCE"):
        return True

    # Drop the long legal disclaimer even when PDFs glue words together
    if "PLEASE" in s_upper and "BOUND" in s_upper and "DUTY" in s_upper:
        return True
    if "PLEASENOTETHATYOUAREBOUNDBYADUTY" in s_compact:
        return True
    if "UNAUTHORISEDDEBITS" in s_compact or "CONCLUSIVELYBINDING" in s_compact:
        return True
    if "CLAIMAGAINSTTHEBANK" in s_compact:
        return True

    # Drop bank footer/address line(s)
    if "UNITED OVERSEAS BANK LIMITED" in s_upper:
        return True
    if "UOB PLAZA" in s_upper or "WWW.UOB.COM.SG" in s_upper:
        return True

    # Drop page indicator
    if re.match(r"^PAGE\s+\d+\s+OF\s+\d+", s_upper):
        return True

    # Drop card product header
    if s_upper.startswith("UOB ONE CARD"):
        return True

    # Drop the masked card number / name continued line
    if "(CONTINUED)" in s_upper:
        return True
    # Typical masked PAN patterns like 1111-1111-1111-1111
    if re.search(r"\b\d{4}-\d{4}-\d{4}-\d{4}\b", s):
        return True

    # Drop "Date Date SGD" variants (sometimes spacing differs)
    if re.fullmatch(r"DATE\s+DATE\s+SGD", s_upper):
        return True
    if "DATEDATESGD" == s_compact:
        return True

    return False


def extract_section_text(pdf_path: Path) -> str | None:
    """Return text between (UOB) transaction table header and SUB TOTAL across all pages."""
    with pdfplumber.open(str(pdf_path)) as pdf:
        all_text = "\n".join((page.extract_text() or "") for page in pdf.pages)
    #print(all_text)
    # Find the first occurrence of the UOB table header, then cut until the first SUB TOTAL after it
    m_start = START_MARKER_RE.search(all_text)
    if not m_start:
        print(f"[SKIP] Start marker not found in: {pdf_path}")
        return None

    after_start = all_text[m_start.start():]

    m_end = END_MARKER_RE.search(after_start)
    if not m_end:
        print(f"[SKIP] End marker not found in: {pdf_path}")
        return None

    return after_start[:m_end.end()]

def clean_lines(section: str) -> list[str]:
    raw = [ln.strip() for ln in section.splitlines()]
    out = []
    for ln in raw:
        if is_redundant_line(ln):
            continue
        # also drop the marker text if it appears merged differently
        if START_MARKER_RE.search(ln):
            continue
        if END_MARKER_RE.search(ln):
            break
        out.append(ln)
    return out

def parse_tx_from_buffer(post_date: str, trans_date: str, buffer_lines: list[str]) -> dict | None:
    """
    buffer_lines contains:
      - first line remainder (after the two dates)
      - plus any continuation lines
    We must produce:
      Transaction Date Captured = trans_date
      Description Captured = description (+ continuation lines, incl Ref No)
      Amount Captured = amount (may be at end of first line, or on its own later line)
    """
    if not buffer_lines:
        return None

    amount = None
    desc_parts = []

    # 1) Check if first line ends with amount
    first = buffer_lines[0].strip()
    m_amt_end = AMT_AT_END_RE.search(first)
    if m_amt_end:
        # If the "amount at end" is really present, split it off from description
        num = m_amt_end.group("num")
        suf = m_amt_end.group("suffix")
        before_amt = first[:m_amt_end.start()].strip()

        # Heuristic: only treat it as an amount if there's actually some description before it
        if before_amt:
            amount = normalize_amount(num, suf)
            if before_amt:
                desc_parts.append(before_amt)
        else:
            desc_parts.append(first)
    else:
        desc_parts.append(first)

    # 2) Process continuation lines:
    #    - if we have not found amount yet, accept amount-only line as amount
    #    - otherwise append into description
    for cont in buffer_lines[1:]:
        s = cont.strip()
        if not s or is_redundant_line(s):
            continue

        if amount is None:
            m_only = AMT_ONLY_RE.match(s)
            if m_only:
                amount = normalize_amount(m_only.group("num"), m_only.group("suffix"))
                continue

        desc_parts.append(s)

    # If amount still not found, we cannot build a valid row
    if amount is None:
        return None

    description = " ".join(desc_parts)
    # Optional: compress multiple spaces
    description = re.sub(r"\s+", " ", description).strip()

    return {
        "Transaction Date Captured": trans_date.strip(),
        "Description Captured": description,
        "Amount Captured": amount,
    }

def get_subfolder_year(pdf_path: Path) -> str:
    """
    Returns the first-level subfolder under base_dir that contains this pdf.
    Example: UOB/2022/xxx.pdf -> "2022"
    If the pdf is directly under UOB/, returns "UOB".
    """
    rel_parts = pdf_path.relative_to(base_dir).parts
    if len(rel_parts) >= 2:
        return rel_parts[0]
    return base_dir.name


def extract_rows_from_pdf(pdf_path: Path) -> list[dict]:
    section = extract_section_text(pdf_path)
    if not section:
        return []

    year_value = get_subfolder_year(pdf_path)
    lines = clean_lines(section)

    rows = []
    current_post = None
    current_trans = None
    buffer_lines = []

    def flush():
        nonlocal current_post, current_trans, buffer_lines
        if current_post and current_trans and buffer_lines:
            rec = parse_tx_from_buffer(current_post, current_trans, buffer_lines)
            if rec:
                rec["Year"] = year_value  # <-- add Year here
                rows.append(rec)
        current_post = None
        current_trans = None
        buffer_lines = []

    i = 0
    while i < len(lines):
        ln = lines[i].strip()

        m = TX_START_RE.match(ln)
        if m:
            flush()
            current_post = m.group("post")
            current_trans = m.group("trans")
            buffer_lines = [m.group("rest").strip()]
        else:
            if current_post is not None:
                buffer_lines.append(ln)
        i += 1

    flush()

    print(f"[OK] {pdf_path} -> {len(rows)} rows")
    return rows


# 4) Run extraction
all_rows = []
for pdf_file in pdf_files:
    all_rows.extend(extract_rows_from_pdf(pdf_file))

# 5) Save to Excel (single sheet, required 3 columns)
df = pd.DataFrame(all_rows, columns=[
    "Transaction Date Captured",
    "Year",
    "Description Captured",
    "Amount Captured",
])

df.to_excel(output_excel, sheet_name="Transactions", index=False)

print(f"Total PDFs scanned: {len(pdf_files)}")
print(f"Total rows extracted: {len(df)}")
print(f"Saved combined output to: {output_excel}")


