from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Dict, List

import numpy as np
import pandas as pd


# =============================
# Paths
# =============================
BASE_DIR = Path(__file__).resolve().parent
#print(BASE_DIR)
ANALYSIS_DIR = BASE_DIR / "AnalysisStatements"
TIKR_DIR = BASE_DIR / "TIKR"  # fallback if AnalysisStatements is empty
MARKETCAP_PATH = BASE_DIR / "MarketCap" / "SGX-MarketCap.xlsx"
OUTPUT_DIR = BASE_DIR / "Output"

SNAPSHOT_YEAR = 2025


# =============================
# Output columns (your spec + Ticker codes)
# =============================
COLUMNS = [
    "Year",  # date headers copied directly from the Excel files (datetime)
    "Revenue",
    "Gross Profit",
    "Gross Margin",
    "Net Income",
    "EPS",
    "Return of Equity",
    "Total Assets",
    "Total Liabilities",
    "Total Equity",
    "Outstanding Shares",
    "Book Value per Share",
    "Free Cashflow",
    "Free Cashflow per share",
    "CashFlow from Operations",  # Cash from Operations
    "Cash",
    "Current Debt",
    "Long Term Debt",
    "Total Debt",
    "Debt-Equity",
    "Dividend per share",
    "Dividend Growth",
    "Dividend yield",
    "Payout Ratio",
    "FCF Dividend Coverage",
    "Ticker codes",
]


# =============================
# Helpers
# =============================
def normalize_label(x: object) -> str:
    s = "" if x is None else str(x)
    s = s.strip().lower()
    s = re.sub("[^a-z0-9]+", " ", s)
    s = " ".join(s.split())
    return s


def parse_ticker_from_filename(filename: str) -> Optional[str]:
    # Expected pattern: TIKR - XXX - Financials ...
    parts = [p.strip() for p in filename.split("-")]
    if len(parts) < 3:
        return None

    if parts[0].strip().upper() != "TIKR":
        return None

    ticker = parts[1].strip().upper()
    return ticker if ticker else None


def safe_filename(s: str, max_len: int = 180) -> str:
    invalid = {"<", ">", ":", '"', "/", "|", "?", "*", chr(92)}
    cleaned = "".join((" " if ch in invalid else ch) for ch in s)
    cleaned = " ".join(cleaned.split())
    return cleaned[:max_len]


def to_timestamp(v: object) -> Optional[pd.Timestamp]:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None

    if isinstance(v, str) and normalize_label(v) in {"ltm", "ttm"}:
        return None

    ts = pd.to_datetime(v, errors="coerce")
    if not pd.isna(ts):
        return ts

    # Excel sometimes exports dates as text like 12/31/06
    ts2 = pd.to_datetime(str(v), format="%m/%d/%y", errors="coerce")
    if not pd.isna(ts2):
        return ts2

    return None


def date_col_map(sheet_df: pd.DataFrame) -> Dict[pd.Timestamp, int]:
    # Header row is row 0, dates begin from column 1
    m: Dict[pd.Timestamp, int] = {}
    for col in range(1, sheet_df.shape[1]):
        ts = to_timestamp(sheet_df.iat[0, col])
        if ts is None:
            continue
        m.setdefault(ts, col)
    return m


def sorted_dates_from_sheet(sheet_df: pd.DataFrame) -> List[pd.Timestamp]:
    return sorted(date_col_map(sheet_df).keys())


def find_row_index_best(sheet_df: pd.DataFrame, target_label: str) -> Optional[int]:
    """
    Row selection rule:
    - If an exact normalized label match exists, ONLY consider the exact matches.
      This avoids accidentally picking a broader "contains" row (e.g. "Long-Term Debt and Capital Leases")
      when the true "Long-Term Debt" row exists.
    - Otherwise, fall back to "contains" matches.
    - If multiple candidates remain, pick the one with the most numeric values across date columns.
    """
    if sheet_df.empty or sheet_df.shape[1] < 2:
        return None

    labels = sheet_df.iloc[:, 0].astype(str).map(normalize_label)
    target = normalize_label(target_label)

    exact = labels[labels == target]
    if len(exact):
        candidates = [int(i) for i in exact.index]
    else:
        contains = labels[labels.str.contains(target, na=False)]
        candidates = []
        for i in contains.index:
            ii = int(i)
            if ii not in candidates:
                candidates.append(ii)

    if not candidates:
        return None

    dmap = date_col_map(sheet_df)
    date_cols = list(dmap.values())
    if not date_cols:
        return candidates[0]

    best_idx = None
    best_score = -1
    for idx in candidates:
        vals = pd.to_numeric(sheet_df.iloc[idx, date_cols], errors="coerce")
        score = int(vals.notna().sum())
        if score > best_score:
            best_score = score
            best_idx = idx

    return best_idx


def extract_values_by_dates(
    sheet_df: pd.DataFrame,
    row_label: str,
    dates: List[pd.Timestamp],
    sheet_name: str,
) -> pd.Series:
    dmap = date_col_map(sheet_df)
    idx = find_row_index_best(sheet_df, row_label)

    if idx is None:
        print(f"Warning: row label not found in {sheet_name}: {row_label}")
        return pd.Series([np.nan] * len(dates), index=dates, dtype="float64")

    out: List[float] = []
    for d in dates:
        if d not in dmap:
            out.append(np.nan)
            continue
        out.append(pd.to_numeric(sheet_df.iat[idx, dmap[d]], errors="coerce"))

    return pd.Series(out, index=dates, dtype="float64")


def lookup_company_name(marketcap_path: Path, ticker: str) -> Optional[str]:
    if not marketcap_path.exists():
        return None

    mdf = pd.read_excel(marketcap_path)
    col_norm = {c: normalize_label(c) for c in mdf.columns}

    ticker_cols = [c for c, n in col_norm.items() if ("ticker" in n or "code" in n or "symbol" in n)]
    name_cols = [c for c, n in col_norm.items() if (("company" in n and "name" in n) or n == "name" or "issuer" in n)]

    if not ticker_cols:
        return None

    tcol = ticker_cols[0]
    match = mdf[mdf[tcol].astype(str).str.upper().str.strip() == ticker.upper()]
    if len(match) == 0:
        return None

    if name_cols:
        return str(match.iloc[0][name_cols[0]]).strip()

    # fallback: pick any other column
    for c in mdf.columns:
        if c != tcol:
            return str(match.iloc[0][c]).strip()

    return None


# =============================
# Build one dataframe per file
# =============================
def build_dataframe_from_file(xlsx_path: Path) -> pd.DataFrame:
    income = pd.read_excel(xlsx_path, sheet_name="Income Statement", header=None)
    bs = pd.read_excel(xlsx_path, sheet_name="Balance Sheet", header=None)
    cf = pd.read_excel(xlsx_path, sheet_name="Cash Flow", header=None)

    # NEW: read Ratios sheet for Dividend yield
    try:
        ratios = pd.read_excel(xlsx_path, sheet_name="Ratios", header=None)
    except Exception:
        ratios = None
        print("Warning: could not read 'Ratios' sheet; Dividend yield will be blank")

    # Copy the date headers directly from the Income Statement
    dates = sorted_dates_from_sheet(income)
    if not dates:
        dates = sorted_dates_from_sheet(bs)

    df = pd.DataFrame({"Year": dates})

    # Income Statement (all dates)
    df["Revenue"] = extract_values_by_dates(income, "Total Revenues", dates, "Income Statement").values
    df["Gross Profit"] = extract_values_by_dates(income, "Gross Profit", dates, "Income Statement").values
    df["Net Income"] = extract_values_by_dates(income, "Net Income to Common", dates, "Income Statement").values
    df["EPS"] = extract_values_by_dates(income, "Normalized Diluted EPS", dates, "Income Statement").values
    df["Outstanding Shares"] = extract_values_by_dates(
        income, "Weighted Average Diluted Shares Outstanding", dates, "Income Statement"
    ).values
    df["Dividend per share"] = extract_values_by_dates(income, "Dividends Per Share", dates, "Income Statement").values

    # Dividend Growth = year-on-year % change in Dividend per share
    dps_series = pd.to_numeric(df["Dividend per share"], errors="coerce")
    df["Dividend Growth"] = dps_series.pct_change()

    # Balance Sheet (all dates)
    df["Total Assets"] = extract_values_by_dates(bs, "Total Assets", dates, "Balance Sheet").values
    df["Total Liabilities"] = extract_values_by_dates(bs, "Total Liabilities", dates, "Balance Sheet").values
    df["Total Equity"] = extract_values_by_dates(bs, "Total Equity", dates, "Balance Sheet").values
    df["Cash"] = extract_values_by_dates(bs, "Total Cash And Short Term Investments", dates, "Balance Sheet").values
    df["Current Debt"] = extract_values_by_dates(bs, "Current Debt", dates, "Balance Sheet").values
    df["Long Term Debt"] = extract_values_by_dates(bs, "Long-Term Debt", dates, "Balance Sheet").values

    # If the source file has no values for these, treat as zero (instead of blank)
    df["Current Debt"] = pd.to_numeric(df["Current Debt"], errors="coerce").fillna(0)
    df["Long Term Debt"] = pd.to_numeric(df["Long Term Debt"], errors="coerce").fillna(0)

    # Cash Flow (all dates)
    df["Free Cashflow"] = extract_values_by_dates(cf, "Free Cash Flow", dates, "Cash Flow").values

    # Derived metrics (all dates)
    df["Gross Margin"] = df["Gross Profit"] / df["Revenue"]
    df["Total Debt"] = df["Current Debt"] + df["Long Term Debt"]
    df["Debt-Equity"] = df["Total Debt"] / df["Total Equity"]
    df["Book Value per Share"] = df["Total Equity"] / df["Outstanding Shares"]
    df["Free Cashflow per share"] = df["Free Cashflow"] / df["Outstanding Shares"]
    
    # Return of Equity = Net Income / Total Equity
    df["Return of Equity"] = np.nan
    ni = pd.to_numeric(df["Net Income"], errors="coerce")
    te = pd.to_numeric(df["Total Equity"], errors="coerce")
    valid_roe = ni.notna() & te.notna() & (te != 0)
    df.loc[valid_roe, "Return of Equity"] = ni[valid_roe] / te[valid_roe]


    # NEW: CFO from Cash Flow tab, row "Cash from Operations"
    df["CashFlow from Operations"] = extract_values_by_dates(cf, "Cash from Operations", dates, "Cash Flow").values
    
    # NEW: Dividend yield from Ratios tab, row "Trailing Dividend Yield"
    if ratios is not None:
        df["Dividend yield"] = extract_values_by_dates(ratios, "Trailing Dividend Yield", dates, "Ratios").values
    else:
        df["Dividend yield"] = np.nan

    # Payout Ratio = Dividend per share / EPS (use dataframe values)
    df["Payout Ratio"] = np.nan
    eps = pd.to_numeric(df["EPS"], errors="coerce")
    dps = pd.to_numeric(df["Dividend per share"], errors="coerce")
    valid = eps.notna() & (eps != 0) & dps.notna()
    df.loc[valid, "Payout Ratio"] = dps[valid] / eps[valid]

    # FCF Dividend Coverage = Free Cashflow per share / Dividend per share
    df["FCF Dividend Coverage"] = np.nan
    fcfps = pd.to_numeric(df["Free Cashflow per share"], errors="coerce")
    valid2 = fcfps.notna() & dps.notna() & (dps != 0)
    df.loc[valid2, "FCF Dividend Coverage"] = (fcfps[valid2] / dps[valid2]) - 1

    # Ensure all expected columns exist
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = np.nan

    df = df[COLUMNS]
    return df


# =============================
# Main batch runner
# =============================
def run_batch() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    files = sorted(ANALYSIS_DIR.glob("*.xlsx"))
    if not files:
        files = sorted(TIKR_DIR.glob("*.xlsx"))

    if not files:
        raise FileNotFoundError(f"No .xlsx files found in {ANALYSIS_DIR} or {TIKR_DIR}.")

    for fp in files:
        print(f"Reading file: {fp.name}")

        ticker = parse_ticker_from_filename(fp.name) or "UNKNOWN"
        company = lookup_company_name(MARKETCAP_PATH, ticker) or "Unknown Company"

        df = build_dataframe_from_file(fp)
        df["Ticker codes"] = ticker

        out_name = safe_filename(f"{ticker} - {company} - {SNAPSHOT_YEAR}.xlsx")
        out_path = OUTPUT_DIR / out_name

        with pd.ExcelWriter(out_path, engine="openpyxl", datetime_format="mm/dd/yy") as writer:
            df.to_excel(writer, index=False, sheet_name="Financials")

        print(f"Saved: {out_path}")


if __name__ == "__main__":
    run_batch()
