"""
╔══════════════════════════════════════════════════════════════╗
║          TATA FUND - EXCEL DATA EXTRACTION BOT               ║
║──────────────────────────────────────────────────────────────║
║  Extracts instrument data from fund portfolio Excel sheets.  ║
║  Supports multiple files & multiple sheets per file.         ║
║  Select specific sheets by NAME, POSITION, or FUND NAME.     ║
║  All data appended into one master output sheet.             ║
║  Enforces uniqueness on: FUND NAME + ISIN (across all).      ║
╚══════════════════════════════════════════════════════════════╝

SETUP:
    pip install pandas openpyxl

USAGE:
    1. Place this script in the same folder as your Excel files.
    2. Edit the CONFIG section below — especially SHEET_SELECTION.
    3. Run:  python excel_extractor_bot.py

OUTPUT COLUMNS:
    FUND NAME | NAME OF THE INSTRUMENT | YIELD (IN %) | RATINGS |
    ISIN CODE | QUANTITY | MKT VAL(Rs. Lacs) | % to NAV
"""

import pandas as pd
import os
import glob


# ══════════════════════════════════════════════════════════════
#  ✏️  CONFIG — Edit these values to match your setup
# ══════════════════════════════════════════════════════════════

# Folder containing input Excel files (use "." for current directory)
INPUT_FOLDER = "./TATA"

# Output file name
OUTPUT_FILE = "Compiled_Tata.xlsx"

# If True: process ALL .xlsx files in INPUT_FOLDER automatically
# If False: only process files listed in SPECIFIC_FILES
PROCESS_ALL_FILES = True

SPECIFIC_FILES = [
    "tata1.xlsx",
    # "tata2.xlsx",
]

# ── SHEET SELECTION ───────────────────────────────────────────
#
# Choose ONE of the three methods below.
# Set the method you want to use, leave the others as-is.
#
# METHOD 1 — By sheet NAME (tab name at the bottom of Excel)
#   Example: ["Sheet1", "TFRSTF", "LiquidFund"]
#
# METHOD 2 — By sheet POSITION (1-based: 1 = first sheet)
#   Example: [1, 3, 5, 7]
#
# METHOD 3 — By FUND NAME (text auto-detected inside the sheet)
#   Example: ["TATA LIQUID FUND", "TATA EQUITY FUND"]
#
# Set SHEET_SELECTION_METHOD to "name", "position", or "fund_name"
# Then fill in SELECTED_SHEETS with your list.
#
SHEET_SELECTION_METHOD = "name"   # ← change to: "name" / "position" / "fund_name"

SELECTED_SHEETS = []
# If SELECTED_SHEETS is empty, ALL sheets in every file are processed.

# On duplicate FUND NAME + ISIN (across all files & sheets):
#   "skip"  → silently skip
#   "log"   → keep first, log duplicates to a separate sheet
DUPLICATE_ACTION = "log"


# ══════════════════════════════════════════════════════════════
#  CORE LOGIC
# ══════════════════════════════════════════════════════════════

TARGET_COLS = [
    "NAME OF THE INSTRUMENT",
    "YIELD ( IN % )",
    "RATINGS",
    "ISIN CODE",
    "QUANTITY",
    "MKT VAL(Rs. Lacs)",
    "% to NAV",
]


def detect_fund_name(df: pd.DataFrame) -> str:
    """Auto-detect fund name from top rows of the sheet."""
    for i in range(20):
        for val in df.iloc[i]:
            s = str(val).strip()
            if "TATA" in s.upper() and "FUND" in s.upper() and len(s) < 80:
                return s.upper()
    return "UNKNOWN FUND"


def find_header_row(df: pd.DataFrame) -> int:
    """Find the row index containing 'NAME OF THE INSTRUMENT'."""
    for i, row in df.iterrows():
        if "NAME OF THE INSTRUMENT" in [str(v).strip() for v in row]:
            return i
    raise ValueError("Could not find header row with 'NAME OF THE INSTRUMENT'.")


def has_market_value(row, mkt_col: str) -> bool:
    """Return True if this row has a valid numeric market value."""
    val = row.get(mkt_col, None)
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    try:
        return float(val) != 0
    except (ValueError, TypeError):
        return False


def get_sheets_to_process(xl: pd.ExcelFile, filepath: str) -> list:
    """
    Return the list of sheet names to process from this file,
    based on SHEET_SELECTION_METHOD and SELECTED_SHEETS config.
    """
    all_sheets = xl.sheet_names

    # If no filter specified → process all sheets
    if not SELECTED_SHEETS:
        return all_sheets

    if SHEET_SELECTION_METHOD == "name":
        selected = []
        for name in SELECTED_SHEETS:
            if name in all_sheets:
                selected.append(name)
            else:
                print(f"      ⚠️  Sheet '{name}' not found in {os.path.basename(filepath)}. Available: {all_sheets}")
        return selected

    elif SHEET_SELECTION_METHOD == "position":
        selected = []
        for pos in SELECTED_SHEETS:
            idx = int(pos) - 1  # convert 1-based to 0-based
            if 0 <= idx < len(all_sheets):
                selected.append(all_sheets[idx])
            else:
                print(f"      ⚠️  Sheet position {pos} out of range. File has {len(all_sheets)} sheets.")
        return selected

    elif SHEET_SELECTION_METHOD == "fund_name":
        selected = []
        for sheet_name in all_sheets:
            df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
            fund = detect_fund_name(df)
            for wanted in SELECTED_SHEETS:
                if wanted.upper() in fund.upper():
                    selected.append(sheet_name)
                    break
        return selected

    else:
        print(f"⚠️  Unknown SHEET_SELECTION_METHOD '{SHEET_SELECTION_METHOD}'. Processing all sheets.")
        return all_sheets


def extract_from_sheet(df: pd.DataFrame, fund_name: str, seen_keys: set) -> tuple:
    """
    Extract all rows that have a market value.
    Returns (extracted_rows, duplicate_rows).
    """
    header_row = find_header_row(df)
    headers = list(df.iloc[header_row])

    data = df.iloc[header_row + 1:].copy()
    data.columns = headers

    mkt_col = "MKT VAL(Rs. Lacs)"
    extracted = []
    duplicates = []

    for _, row in data.iterrows():
        name = str(row.get("NAME OF THE INSTRUMENT", "")).strip()

        if not name or name == "nan":
            continue

        # CORE RULE: only copy rows that have a market value
        if not has_market_value(row, mkt_col):
            continue

        isin = str(row.get("ISIN CODE", "")).strip()
        is_valid_isin = bool(isin and isin != "nan" and len(isin) > 5)
        clean_isin = isin if is_valid_isin else ""

        # Uniqueness key across ALL files and sheets
        key = (fund_name, clean_isin if clean_isin else f"__NO_ISIN__{name}")

        if key in seen_keys:
            duplicates.append({
                "FUND NAME": fund_name,
                "NAME OF THE INSTRUMENT": name,
                "ISIN CODE": clean_isin,
                "REASON": "Duplicate FUND NAME + ISIN",
            })
            continue

        seen_keys.add(key)

        extracted.append({
            "FUND NAME": fund_name,
            "NAME OF THE INSTRUMENT": name,
            "YIELD ( IN % )": row.get("YIELD ( IN % )", ""),
            "RATINGS": row.get("RATINGS", ""),
            "ISIN CODE": clean_isin,
            "QUANTITY": row.get("QUANTITY", ""),
            "MKT VAL(Rs. Lacs)": row.get(mkt_col, ""),
            "% to NAV": row.get("% to NAV", ""),
        })

    return extracted, duplicates


def process_file(filepath: str, seen_keys: set) -> tuple:
    """Process selected sheets in a single Excel file."""
    print(f"\n📂 Processing: {os.path.basename(filepath)}")
    xl = pd.ExcelFile(filepath)

    sheets_to_process = get_sheets_to_process(xl, filepath)
    print(f"   Sheets to process ({len(sheets_to_process)}): {sheets_to_process}")

    all_extracted = []
    all_duplicates = []

    for sheet_name in sheets_to_process:
        print(f"   └─ Sheet: '{sheet_name}'")
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        fund_name = detect_fund_name(df)
        print(f"      Fund Name : {fund_name}")

        try:
            extracted, duplicates = extract_from_sheet(df, fund_name, seen_keys)
            print(f"      ✅ Extracted {len(extracted)} rows | Duplicates skipped: {len(duplicates)}")
            all_extracted.extend(extracted)
            all_duplicates.extend(duplicates)
        except ValueError as e:
            print(f"      ⚠️  Skipped — {e}")

    return all_extracted, all_duplicates


def save_output(extracted: list, duplicates: list, output_path: str):
    """Save all extracted data to one master sheet in the output Excel."""
    out_df = pd.DataFrame(extracted, columns=["FUND NAME"] + TARGET_COLS)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Extracted Data", index=False)

        if DUPLICATE_ACTION == "log" and duplicates:
            dup_df = pd.DataFrame(duplicates)
            dup_df.to_excel(writer, sheet_name="Duplicates Log", index=False)
            print(f"\n📋 {len(duplicates)} duplicate(s) logged → 'Duplicates Log' sheet")

    print(f"\n✅ Output saved → {output_path}")
    print(f"   Total rows extracted : {len(extracted)}")
    print(f"   Unique funds found   : {out_df['FUND NAME'].nunique()}")
    print(f"   Funds: {list(out_df['FUND NAME'].unique())}")


def main():
    if PROCESS_ALL_FILES:
        files = sorted(glob.glob(os.path.join(INPUT_FOLDER, "*.xlsx")))
        files = [f for f in files if os.path.basename(f) != OUTPUT_FILE]
    else:
        files = [os.path.join(INPUT_FOLDER, f) for f in SPECIFIC_FILES]

    if not files:
        print("⚠️  No Excel files found. Check INPUT_FOLDER or SPECIFIC_FILES in CONFIG.")
        return

    print(f"🔍 Found {len(files)} file(s) to process:")
    for f in files:
        print(f"   • {os.path.basename(f)}")

    all_extracted = []
    all_duplicates = []
    seen_keys = set()  # Global uniqueness across ALL files and ALL sheets

    for filepath in files:
        if not os.path.exists(filepath):
            print(f"❌ File not found: {filepath}")
            continue
        extracted, duplicates = process_file(filepath, seen_keys)
        all_extracted.extend(extracted)
        all_duplicates.extend(duplicates)

    if not all_extracted:
        print("\n⚠️  No data extracted. Please check your files and sheet structure.")
        return

    save_output(all_extracted, all_duplicates, OUTPUT_FILE)


if __name__ == "__main__":
    main()
