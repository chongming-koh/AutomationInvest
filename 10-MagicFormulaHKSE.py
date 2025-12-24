import pandas as pd
import re
from pathlib import Path

# ---------- Helpers ----------
def extract_ticker_from_filename(name: str) -> str:
    m = re.search(r"TIKR\s*-\s*([A-Z0-9]+)\s*-\s*Financials", name, flags=re.IGNORECASE)
    if not m:
        raise ValueError(f"Could not extract ticker from filename: {name}")
    return m.group(1).upper()

# UPDATED: Some TIKR filenames end with a 3-letter currency code, e.g. '-USD.xlsx' or '-CNY.xlsx'
def extract_ccy_from_filename(name: str) -> str | None:
    m = re.search(
        r"[\u2010\u2011\u2012\u2013\u2014\u2212\-]\s*([A-Z]{3})\s*\.xlsx$",
        name,
        flags=re.IGNORECASE,
    )
    if not m:
        return None
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
MARKETCAP_PATH = SCRIPT_DIR / "MarketCap" / "HKSE-MarketCap.xlsx"
OUTPUT_PATH = SCRIPT_DIR / "HKSE-Magic-Formula.xlsx"

if not TIKR_DIR.exists():
    raise FileNotFoundError(f"TIKR folder not found: {TIKR_DIR}")

if not MARKETCAP_PATH.exists():
    raise FileNotFoundError(f"HKSE-MarketCap.xlsx not found: {MARKETCAP_PATH}")

# ---------- Load Market Cap file ----------
mc = pd.read_excel(MARKETCAP_PATH)

# Update for CCY calculation: HKD -> CCY FX rates (multiplier)
# - Keep these rates here so you can update them easily when needed.
HKD_TO_CCY_RATES: dict[str, float] = {
    "CNY": 0.90,
    "USD": 0.13,
    "SGD": 0.16,
    "EUR": 0.11,
    "JPY": 20.19,
}


def market_cap_HKD_to_ccy_millions(mkt_cap_HKD: float | int | None, ccy: str | None) -> float | None:
    """Convert a HKD market cap value to CCY (if provided) and return in millions."""
    if mkt_cap_HKD is None:
        return None

    amt = pd.to_numeric(mkt_cap_HKD, errors="coerce")
    if pd.isna(amt):
        return None

        # Update for CCY calculation: if CCY is empty/NaN, keep market cap in HKD.
    if ccy is None or pd.isna(ccy) or str(ccy).strip() == "":
        # No CCY suffix in filename (or missing CCY), keep market cap in HKD.
        return float(amt) / 1_000_000

    ccy_u = str(ccy).strip().upper()
    if ccy_u not in HKD_TO_CCY_RATES:
        msg = (
            f"Update for CCY calculation: Missing HKD-> {ccy_u} rate. "
            "Please update HKD_TO_CCY_RATES in the script."
        )
        print(msg)
        raise ValueError(msg)

    return (float(amt) * HKD_TO_CCY_RATES[ccy_u]) / 1_000_000

ticker_col = next((c for c in mc.columns if _norm(c) in {"ticker", "tickers"}), None)
company_col = next((c for c in mc.columns if _norm(c) in {"company name", "company"}), None)
if ticker_col is None or company_col is None:
    raise ValueError("Could not find required columns 'Ticker' and 'Company Name' in SGX-MarketCap.xlsx.")

# UPDATED: Market cap source file changed - use the "Mkt Cap" column (instead of the last column)
mkt_cap_col = next((c for c in mc.columns if _norm(c) in {"mkt cap", "market cap", "market capitalization"}), None)
if mkt_cap_col is None:
    raise ValueError("Could not find required column 'Mkt Cap' in SGX-MarketCap.xlsx.")

marketcap_value_col = mkt_cap_col

# Update for CCY calculation: keep Market Cap as raw HKD in the lookup.
# Conversion to CCY (and to millions) is done per-row based on the file's CCY suffix.
mc[marketcap_value_col] = pd.to_numeric(mc[marketcap_value_col], errors="coerce")
mc_lookup = mc.copy()
mc_lookup[ticker_col] = mc_lookup[ticker_col].astype(str).str.strip().str.upper()
mc_lookup = mc_lookup.set_index(ticker_col)

# ---------- Process TIKR financial files ----------
#tikr_files = sorted(TIKR_DIR.glob("TIKR - * - Financials (*.xlsx"))
tikr_files = sorted(TIKR_DIR.glob("TIKR - * - Financials*.xlsx"))

print("Picked up these TIKR files:") # UPDATED
for p in tikr_files: # UPDATED
    print(" -", p.name) # UPDATED

if not tikr_files:
    raise FileNotFoundError(f"No TIKR financial files found in: {TIKR_DIR}")

rows = []
for fpath in tikr_files:
    ticker = extract_ticker_from_filename(fpath.name)

    # UPDATED: Extract optional 3-letter currency code suffix from filename (if present)
    ccy = extract_ccy_from_filename(fpath.name)    
    company_name = None
    market_cap = None
    raw_market_cap_HKD = None
    if ticker in mc_lookup.index:
        company_name = mc_lookup.at[ticker, company_col]
        raw_market_cap_HKD = mc_lookup.at[ticker, marketcap_value_col]

    if isinstance(company_name, pd.Series):
        company_name = company_name.iloc[0]
    if isinstance(raw_market_cap_HKD, pd.Series):
        raw_market_cap_HKD = raw_market_cap_HKD.iloc[0]

    # Update for CCY calculation: convert raw HKD market cap into the row CCY (if any), then to millions.
    if raw_market_cap_HKD is not None:
        market_cap = market_cap_HKD_to_ccy_millions(raw_market_cap_HKD, ccy)

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
        "CCY": ccy,  # UPDATED
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
        "Rank_Earn_Yield": None,  # Update for Ranking
        "Rank_ROC": None,  # Update for Ranking
        "Overall_Rank": None,  # Update for Ranking,
    })

df = pd.DataFrame(
    rows,
    columns=[
        "Ticker codes",
        "Company Name",
        "CCY",  # UPDATED
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
        "Rank_Earn_Yield",  # Update for Ranking
        "Rank_ROC",  # Update for Ranking
        "Overall_Rank",  # Update for Ranking,
    ],
)

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

# Ranking code updates: Rank Earning Yield and Return on Capital, then compute Overall Rank
# Ranking code updates: For non-positive values, use NaN in Rank columns (instead of text)

# Ensure numeric types for ranking
# (prevents ranking issues if values are stored as strings)
df["Earning Yield"] = pd.to_numeric(df["Earning Yield"], errors="coerce").fillna(0)
df["Return on Capital"] = pd.to_numeric(df["Return on Capital"], errors="coerce").fillna(0)

# Rank_Earn_Yield: only rank values > 0, else keep NaN
_mask_ey = df["Earning Yield"] > 0
_ey_ranks = df.loc[_mask_ey, "Earning Yield"].rank(method="dense", ascending=False)
df["Rank_Earn_Yield"] = pd.Series(pd.NA, index=df.index, dtype="Int64")
df.loc[_mask_ey, "Rank_Earn_Yield"] = _ey_ranks.astype("Int64").to_numpy()

# Rank_ROC: only rank values > 0, else keep NaN
_mask_roc = df["Return on Capital"] > 0
_roc_ranks = df.loc[_mask_roc, "Return on Capital"].rank(method="dense", ascending=False)
df["Rank_ROC"] = pd.Series(pd.NA, index=df.index, dtype="Int64")
df.loc[_mask_roc, "Rank_ROC"] = _roc_ranks.astype("Int64").to_numpy()

# Overall_Rank: sum of the two ranks (will be NaN if either is NaN)
df["Overall_Rank"] = (df["Rank_Earn_Yield"] + df["Rank_ROC"]).astype("Int64")

# Sort by Overall_Rank (smallest first). NaN goes to the bottom.
df = df.sort_values(by=["Overall_Rank"], ascending=True, na_position="last")

df.to_excel(OUTPUT_PATH, index=False)
print(f"Saved output to: {OUTPUT_PATH}")
