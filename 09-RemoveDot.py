"""
This code is to update the file names from TIKR - 590 - Financials (3.31.06 - 3.31.25)-HKD..xlsx to TIKR - 590 - Financials (3.31.06 - 3.31.25)-HKD.xlsx 
When i save the file name, i save with additional "." which will cause renaming of file name with ccy with errors.
Run this code first before running the ConvertFileNames.py
"""

from pathlib import Path
import re

FOLDER = Path("MagicFormula\TIKR")   # change if needed
DRY_RUN = False                # set to False to actually rename

# Only fix the specific pattern "-HKD..xlsx" (or any currency "-XXX..xlsx")
# so files like "-HKD.xlsx" remain unchanged.
EXTRA_DOT_AFTER_CCY_RE = re.compile(r"(-[A-Za-z]{3})\.\.(xlsx)$", flags=re.IGNORECASE)

def safe_rename(src: Path, dst: Path) -> None:
    if src == dst:
        return

    candidate = dst
    i = 1
    while candidate.exists():
        candidate = dst.with_name(f"{dst.stem} ({i}){dst.suffix}")
        i += 1

    if DRY_RUN:
        print(f"[DRY RUN] {src.name}  ->  {candidate.name}")
    else:
        src.rename(candidate)
        print(f"Renamed: {src.name}  ->  {candidate.name}")

def main():
    if not FOLDER.exists() or not FOLDER.is_dir():
        raise FileNotFoundError(f"Folder not found: {FOLDER.resolve()}")

    for f in FOLDER.iterdir():
        if not f.is_file():
            continue
        if f.suffix.lower() != ".xlsx":
            continue

        old_name = f.name

        # Turn "-HKD..xlsx" into "-HKD.xlsx"
        new_name = EXTRA_DOT_AFTER_CCY_RE.sub(r"\1.\2", old_name)

        if new_name != old_name:
            safe_rename(f, f.with_name(new_name))
        else:
            if DRY_RUN:
                print(f"[DRY RUN] (no change) {old_name}")

if __name__ == "__main__":
    main()
