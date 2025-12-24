'''
The purpose of this code is to change the file name. For filename without any currency code, update the file name with "CNY"
As for file name that ends with HKD, remove it.
'''

from pathlib import Path
import re

FOLDER = Path("MagicFormula\TIKR")   # change if needed
DRY_RUN = False                # set to False to actually rename

# Matches a 3-letter currency code at the very end of the stem, like "-HKD"
CURRENCY_AT_END_RE = re.compile(r"-[A-Za-z]{3}$")

def safe_rename(src: Path, dst: Path) -> None:
    """Rename src -> dst, avoiding collisions by appending (1), (2), ... if needed."""
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

        stem = f.stem  # filename without extension

        # Rule 1: remove trailing "-HKD"
        if stem.lower().endswith("-hkd"):
            new_stem = stem[:-4]  # remove last 4 chars: "-HKD"
            new_name = f"{new_stem}.xlsx"
            safe_rename(f, f.with_name(new_name))
            continue

        # Rule 2: if no trailing "-CCC" currency, append "-CNY"
        if CURRENCY_AT_END_RE.search(stem) is None:
            new_name = f"{stem}-CNY.xlsx"
            safe_rename(f, f.with_name(new_name))
            continue

        # Otherwise, leave as-is
        if DRY_RUN:
            print(f"[DRY RUN] (no change) {f.name}")

if __name__ == "__main__":
    main()
