import pandas as pd
import re
from pathlib import Path

# ---------- Helpers ----------
def extract_ticker_from_filename(name: str) -> str:
    m = re.search(r"TIKR\s*-\s*([A-Z0-9]+)\s*-\s*Financials", name, flags=re.IGNORECASE)
    if not m:
        raise ValueError(f"Could not extract ticker from filename: {name}")
    return m.group(1).upper()

def _norm(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r"[\u2010\u2011\u2012\u2013\u2014\u2212\-]", " ", s)  # hyphen variants
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def pick_year_column(sheet_df: pd.DataFrame, preferred_year: int = 2025) -> tuple[int, int]:
    header = sheet_df.iloc[0, 1:]  # row 1, col B onward

    # TIKR headers are mm/dd/yy
    dates = pd.to_datetime(header, errors="coerce", format="%m/%d/%y")
    years = dates.dt.year

    valid = years.notna()
    if not valid.any():
        raise ValueError("No parseable date headers found in row 1 starting from column B.")

    year_to_cols = {}
    for offset, (ok, y) in enumerate(zip(valid.tolist(), years.tolist()), start=1):
        if ok:
            year_to_cols.setdefault(int(y), []).append(offset)

    if preferred_year in year_to_cols:
        chosen_year = preferred_year
        chosen_col = year_to_cols[preferred_year][-1]
    else:
        chosen_year = max(year_to_cols.keys())
        chosen_col = year_to_cols[chosen_year][-1]

    return chosen_year, chosen_col


def find_row_value(sheet_df: pd.DataFrame, label_variants: list[str], value_col: int) -> float | None:
    col0 = sheet_df.iloc[:, 0].astype(str).map(_norm)
    variants = [_norm(v) for v in label_variants]

    hit = None
    for v in variants:
        m = col0.str.contains(re.escape(v))
        idxs = sheet_df.index[m].tolist()
        if idxs:
            hit = idxs[0]
            break

    if hit is None:
        return None

    raw = sheet_df.iat[hit, value_col]
    val = pd.to_numeric(raw, errors="coerce")
    return None if pd.isna(val) else float(val)

# ---------- Paths (relative to this .py file) ----------
SCRIPT_DIR = Path(__file__).resolve().parent

TIKR_DIR = SCRIPT_DIR / "TIKR"
MARKETCAP_PATH = SCRIPT_DIR / "MarketCap" / "SGX-MarketCap.xlsx"
OUTPUT_PATH = SCRIPT_DIR / "SGX-Magic-Formula.xlsx"

if not TIKR_DIR.exists():
    raise FileNotFoundError(f"TIKR folder not found: {TIKR_DIR}")

if not MARKETCAP_PATH.exists():
    raise FileNotFoundError(f"SGX-MarketCap.xlsx not found: {MARKETCAP_PATH}")

# ---------- Load Market Cap file ----------
mc = pd.read_excel(MARKETCAP_PATH)
ticker_col = next((c for c in mc.columns if _norm(c) in {"ticker", "tickers"}), None)
company_col = next((c for c in mc.columns if _norm(c) in {"company name", "company"}), None)
if ticker_col is None or company_col is None:
    raise ValueError("Could not find required columns 'Ticker' and 'Company Name' in SGX-MarketCap.xlsx.")

# UPDATED: Market cap source file changed - use the "Mkt Cap" column (instead of the last column)
mkt_cap_col = next((c for c in mc.columns if _norm(c) in {"mkt cap", "market cap", "market capitalization"}), None)
if mkt_cap_col is None:
    raise ValueError("Could not find required column 'Mkt Cap' in SGX-MarketCap.xlsx.")

marketcap_value_col = mkt_cap_col

# UPDATED: Convert "Mkt Cap" column to millions once at source
mc[marketcap_value_col] = pd.to_numeric(mc[marketcap_value_col], errors="coerce").fillna(0) / 1_000_000
mc_lookup = mc.copy()
mc_lookup[ticker_col] = mc_lookup[ticker_col].astype(str).str.strip().str.upper()
mc_lookup = mc_lookup.set_index(ticker_col)

# ---------- Process TIKR financial files ----------
tikr_files = sorted(TIKR_DIR.glob("TIKR - * - Financials (*.xlsx"))

print("Picked up these TIKR files:") # UPDATED
for p in tikr_files: # UPDATED
    print(" -", p.name) # UPDATED

if not tikr_files:
    raise FileNotFoundError(f"No TIKR financial files found in: {TIKR_DIR}")

rows = []
for fpath in tikr_files:
    ticker = extract_ticker_from_filename(fpath.name)

    company_name = None
    market_cap = None
    if ticker in mc_lookup.index:
        company_name = mc_lookup.at[ticker, company_col]
        market_cap = mc_lookup.at[ticker, marketcap_value_col]
    if isinstance(company_name, pd.Series):
        company_name = company_name.iloc[0]
    if isinstance(market_cap, pd.Series):
        market_cap = market_cap.iloc[0]

    # Income Statement
    inc = pd.read_excel(fpath, sheet_name="Income Statement", header=None)
    _, col_inc = pick_year_column(inc, preferred_year=2025) #update the year if i want to extract past years. If cannot find, default to most recent year
    operating_income = find_row_value(inc, ["Operating Income"], col_inc)

    # Balance Sheet
    bs = pd.read_excel(fpath, sheet_name="Balance Sheet", header=None)
    _, col_bs = pick_year_column(bs, preferred_year=2025) #update the year if i want to extract past years. If cannot find, default to most recent year

    long_term_debt = find_row_value(bs, ["Long-Term Debt", "Long Term Debt"], col_bs)
    total_current_assets = find_row_value(bs, ["Total Current Assets"], col_bs)
    total_current_liabilities = find_row_value(bs, ["Total Current Liabilities"], col_bs)
    net_ppe = find_row_value(
        bs,
        ["Net Property Plant And Equipment", "Net Property Plant Equipment"],
        col_bs
    )
    current_debt = find_row_value(bs, ["Current Debt"], col_bs)

    rows.append({
        "Ticker codes": ticker,
        "Company Name": company_name,
        "Operating Income": operating_income,
        "Market Cap": market_cap,
        "Current Debt": current_debt,
        "Long Term Debt": long_term_debt,
        "Total Current Assets": total_current_assets,
        "Total Current Liabilities": total_current_liabilities,
        "Enterprise Value": None,
        "Net Property Plant Equipment": net_ppe,
        "Earning Yield": None,
        "Return on Capital": None,
    })

df = pd.DataFrame(rows, columns=[
    "Ticker codes",
    "Company Name",
    "Operating Income",
    "Market Cap",
    "Current Debt",
    "Long Term Debt",
    "Total Current Assets",
    "Total Current Liabilities",
    "Enterprise Value",
    "Net Property Plant Equipment",
    "Earning Yield",
    "Return on Capital",
])

# UPDATED (Step 1-5): calculations for Enterprise Value, Earning Yield, Return on Capital
# Step 2: fill empty numeric cells with 0 to avoid calculation issues
_numeric_cols = [
    "Operating Income",
    "Market Cap",
    "Current Debt",
    "Long Term Debt",
    "Total Current Assets",
    "Total Current Liabilities",
    "Net Property Plant Equipment",
]

df[_numeric_cols] = df[_numeric_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

# Step 3: Enterprise Value = (Market Cap + Current Debt + Long Term Debt) - (Total Current Assets - Total Current Liabilities)
df["Enterprise Value"] = (
    (df["Market Cap"] + df["Current Debt"] + df["Long Term Debt"]) -
    (df["Total Current Assets"] - df["Total Current Liabilities"])
)

# Step 4: Earning Yield = Operating Income / Enterprise Value
_ev_denom = df["Enterprise Value"].replace(0, pd.NA)
df["Earning Yield"] = (df["Operating Income"] / _ev_denom).fillna(0)

# Step 5: Return on Capital = Operating Income / (Net PPE + (Total Current Assets - Total Current Liabilities))
_roc_denom = (
    df["Net Property Plant Equipment"] +
    (df["Total Current Assets"] - df["Total Current Liabilities"])
).replace(0, pd.NA)

df["Return on Capital"] = (df["Operating Income"] / _roc_denom).fillna(0)

df.to_excel(OUTPUT_PATH, index=False)
print(f"Saved output to: {OUTPUT_PATH}")
