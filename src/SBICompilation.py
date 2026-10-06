import openpyxl
from openpyxl import Workbook
import re
import datetime

LOG_FILE = "bank6_log.txt"

TARGETS = {
    "isin": "isin",
    "treps": "f) treps / reverse repo investments",
    "treps_end": "total",
    "net_receivables": "net receivable / payable",
    "stop_key": "grand total (aum)"
}


# =========================
# LOGGING
# =========================

def log(message):
    t = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    safe = message.encode("ascii", "replace").decode("ascii")
    print(f"{t} {safe}")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{t} {safe}\n")


# =========================
# CLEAN TEXT
# =========================

def clean(v):
    if v is None:
        return ""
    v = str(v).lower()
    v = v.replace("\xa0", " ").replace("\n", " ").replace("\t", " ")
    v = re.sub(r"\s+", " ", v)
    return v.strip()


# =========================
# PROCESS ONE SHEET
# =========================

def process_sheet(ws, out_ws):

    # Header from D2
    header = ws["D3"].value
    log(f"Header extracted: {header}")

    max_row = ws.max_row
    col_count = ws.max_column

    r = 7     # <<<<<<<<<<<<<< START HERE NOW

    while r <= max_row:

        colC = clean(ws[f"C{r}"].value)
        colD = clean(ws[f"D{r}"].value)

        # STOP IF GRAND TOTAL (AUM)
        if colC == TARGETS["stop_key"]:
            log(f"Row {r}: Found GRAND TOTAL (AUM) → STOP SHEET")
            break

        # Inject header into Column A
        ws[f"A{r}"].value = header

        # ==================================================
        # RULE 1: COLUMN D HAS DATA → DIRECT ISIN PROCESSING
        # ==================================================
        if colD != "":
            if colD == TARGETS["isin"]:
                log(f"Row {r}: D = ISIN (header) → SKIPPED")
                r += 1
                continue

            # Valid ISIN row → COPY
            out_ws.append([
                ws.cell(row=r, column=c).value
                for c in range(1, col_count + 1)
            ])
            log(f"Row {r}: ISIN data → COPIED")
            r += 1
            continue

        # ==================================================
        # RULE 2: COLUMN D BLANK → APPLY COLUMN C CHECKS
        # ==================================================

        # Net Receivables
        if colC == TARGETS["net_receivables"]:
            out_ws.append([
                ws.cell(row=r, column=c).value
                for c in range(1, col_count + 1)
            ])
            log(f"Row {r}: Net Receivable / Payable → COPIED")
            r += 1
            continue

        # TREPS block logic
        if colC == TARGETS["treps"]:
            log(f"Row {r}: TREPS trigger → BEGIN TREPS BLOCK")
            r += 1

            while r <= max_row:
                nextC = clean(ws[f"C{r}"].value)

                if nextC == TARGETS["treps_end"]:
                    log(f"Row {r}: FOUND Total → END TREPS BLOCK")
                    break

                out_ws.append([
                    ws.cell(row=r, column=c).value
                    for c in range(1, col_count + 1)
                ])
                log(f"Row {r}: TREPS block row → COPIED")
                r += 1

            continue

        # Everything else → Skip
        log(f"Row {r}: No match → SKIPPED")
        r += 1


# =========================
# PROCESS FULL WORKBOOK
# =========================

def process_file(input_path, output_path):
    open(LOG_FILE, "w").close()
    log("=== BANK #6 SCRIPT STARTED ===")

    wb = openpyxl.load_workbook(input_path, data_only=True)

    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "Compiled"

    for sheet_name in wb.sheetnames:
        log(f"\n--- Processing Sheet: {sheet_name} ---")
        ws = wb[sheet_name]
        process_sheet(ws, out_ws)

    out_wb.save(output_path)
    log(f"\nOUTPUT SAVED: {output_path}")
    log("=== BANK #6 SCRIPT COMPLETE ===")


# =========================
# RUN
# =========================

if __name__ == "__main__":
    process_file(
        input_path="SBI.xlsx",
        output_path="./CompiledFiles/Compiled_SBI.xlsx"
    )
