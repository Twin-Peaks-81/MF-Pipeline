import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from rapidfuzz import fuzz
import traceback
import datetime
import re

# =========================
# TARGET PHRASES
# =========================
TARGETS = {
    "grand_total": "grand total",
    "subtotal": "sub total",
    "isin": "isin",
    "net_receivables": "net receivables / (payables)",
    "treps": "treps / reverse repo"
}

STRICT_FUZZ = 85   # For SUB TOTAL & ISIN
SOFT_FUZZ = 70     # For TREPS & Net Receivables
LOG_FILE = "script_log.txt"


# =========================
# LOGGING
# =========================
def log(message: str) -> None:
    timestamp = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    safe = message.encode("ascii", "replace").decode("ascii")
    line = f"{timestamp} {safe}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# =========================
# CLEAN TEXT
# =========================
def clean(value) -> str:
    """Normalize text: lowercase, remove NBSP, collapse spaces."""
    if value is None:
        return ""
    text = str(value).lower()
    text = text.replace("\xa0", " ").replace("\n", " ").replace("\t", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


# =========================
# FUZZY MATCHER
# =========================
def fuzzy_match(value, key: str) -> bool:
    cleaned = clean(value)
    target = TARGETS[key]

    threshold = SOFT_FUZZ if key in ("treps", "net_receivables") else STRICT_FUZZ
    score = fuzz.partial_ratio(cleaned, target)

    log(f"      fuzzy-check: '{cleaned}' vs '{target}' | score={score} threshold={threshold}")
    return score >= threshold


# =========================
# EXACT MATCH FOR GRAND TOTAL
# =========================
def exact_grand_total(value) -> bool:
    """Match 'GRAND TOTAL', allowing trailing text like '(AUM)'
    (ignoring case & extra spaces)."""
    cleaned = clean(value)
    match = cleaned.startswith(TARGETS["grand_total"])
    log(f"      exact-check: '{cleaned}' starts with 'grand total'? -> {match}")
    return match


# =========================
# COLUMN OFFSET HANDLING
# =========================
# Some fortnights, column B is left blank and every column the script
# normally reads (B, C, J, ...) shifts one to the right. Instead of
# hard-coding B/C/J, we detect per-sheet whether column B is empty while
# the "next" column has data, and shift our column references accordingly.
def col(letter: str, offset: int) -> str:
    """Return the column letter shifted by `offset` positions to the right."""
    idx = column_index_from_string(letter) + offset
    return get_column_letter(idx)


# NOTE: verified against real AdityaBirla.xlsx (this fortnight) and oldAb.xlsx
# (last fortnight). Two things changed this fortnight, both handled below:
#   1. Every sheet shifted one column right (offset=1) -- column B is no
#      longer blank, it now holds an internal security-code column, so we
#      can't detect the shift by checking "is B empty". Instead we locate
#      the literal "ISIN" column header, which is a stable anchor.
#   2. "Grand Total" is now written as "Grand Total (AUM)" -- exact_grand_total
#      below uses a startswith check instead of exact equality to handle this.
def detect_offset(ws, header_scan_rows: int = 20, max_offset: int = 5) -> int:
    """
    Detect how many columns everything has shifted right by.

    Heuristic: the header row always has a literal 'ISIN' column heading.
    In the original layout that heading sits in column C (index 3). If the
    whole sheet has shifted right by N columns, the ISIN heading will be
    found in column (3 + N) instead. This is far more reliable than
    checking whether column B is empty, since column B can be repurposed
    (e.g. filled with internal security codes) rather than left blank.
    """
    for r in range(1, min(ws.max_row, header_scan_rows) + 1):
        for c in range(1, ws.max_column + 1):
            if clean(ws.cell(row=r, column=c).value) == "isin":
                offset = c - 3  # 3 = original column index of "C"
                if offset >= 0:
                    log(f"Detected column offset = {offset} (ISIN header found in column {get_column_letter(c)}, row {r})")
                    return offset

    log("Could not detect column offset from ISIN header; defaulting to 0")
    return 0


# =========================
# MAIN PROCESSOR
# =========================
def process_file(input_path: str, output_path: str) -> None:
    open(LOG_FILE, "w").close()
    log("========== SCRIPT STARTED ==========")

    # Load Excel
    try:
        wb = openpyxl.load_workbook(input_path, data_only=True)
        log(f"Loaded workbook: {input_path}")
    except Exception as e:
        log(f"ERROR loading workbook: {e}")
        log(traceback.format_exc())
        return

    # Create output workbook
    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "Compiled"

    # Process each sheet
    for sheet_name in wb.sheetnames:
        log("")
        log(f"--- Processing Sheet: {sheet_name} ---")
        ws = wb[sheet_name]

        # ==========================================
        # DETECT COLUMN OFFSET FOR THIS SHEET
        # ==========================================
        offset = detect_offset(ws)
        B, C, J = col("B", offset), col("C", offset), col("J", offset)
        log(f"Using columns -> B:{B} C:{C} J:{J} (offset={offset})")

        # ==========================================
        # FIND GRAND TOTAL (exact match only)
        # ==========================================
        grand_total_row = None

        for r in range(1, ws.max_row + 1):
            if exact_grand_total(ws[f"{B}{r}"].value):
                grand_total_row = r
                log(f"Found EXACT GRAND TOTAL at row {r}")
                break

        if grand_total_row is None:
            log("No EXACT GRAND TOTAL found. Skipping sheet.")
            continue

        # ==========================================
        # Fill column J with B1
        # ==========================================
        header = ws[f"{B}1"].value
        log(f"{B}1 header to fill in {J} column: {header}")

        for r in range(2, grand_total_row + 1):
            ws[f"{J}{r}"].value = header

        # ==========================================
        # Row scanning
        # ==========================================
        r = 2
        while r <= grand_total_row:

            colB_raw = ws[f"{B}{r}"].value
            colC_raw = ws[f"{C}{r}"].value
            colB = clean(colB_raw)
            colC = clean(colC_raw)

            log(f"Row {r}: B='{colB}' C='{colC}'")

            # -----------------------------
            # Case A: Column C has data
            # -----------------------------
            if colC != "":
                if fuzzy_match(colC, "isin"):
                    log(f"Row {r}: ISIN row -> skipped")
                    r += 1
                    continue

                out_ws.append([ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)])
                log(f"Row {r}: Copied (C has data)")
                r += 1
                continue

            # -----------------------------
            # Case B: Column C blank
            # -----------------------------
            if fuzzy_match(colB, "net_receivables"):
                out_ws.append([ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)])
                log(f"Row {r}: Copied (Net Receivables)")
                r += 1
                continue

            if fuzzy_match(colB, "treps"):
                log(f"Row {r}: TREPS detected -> scanning block")
                r += 1  # skip the TREPS row

                while r <= grand_total_row:
                    b_clean = clean(ws[f"{B}{r}"].value)

                    # STOP at SUB TOTAL (strict fuzzy)
                    if fuzzy_match(b_clean, "subtotal"):
                        log(f"Row {r}: SUB TOTAL found -> end TREPS block")
                        break

                    out_ws.append([
                        ws.cell(row=r, column=c).value
                        for c in range(1, ws.max_column + 1)
                    ])
                    log(f"Row {r}: Copied (TREPS block)")
                    r += 1

                continue

            r += 1  # default step

        log(f"Finished sheet: {sheet_name}")

    # Save output file
    try:
        out_wb.save(output_path)
        log(f"Saved output -> {output_path}")
    except Exception as e:
        log(f"ERROR saving output: {e}")
        log(traceback.format_exc())

    log("========== SCRIPT COMPLETE ==========")


# =========================
# RUN SCRIPT
# =========================
if __name__ == "__main__":
    process_file(
        input_path="AdityaBirla.xlsx",
        output_path="./CompiledFiles/Compiled_AdityaBirla.xlsx"
    )