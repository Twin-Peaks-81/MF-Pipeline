import pandas as pd
import re
import os

# ==============================================================================
# OUTPUT COLUMNS
# ==============================================================================
OUTPUT_COLUMNS = [
    "Fund Name", "Security Type", "Name of the Instrument", "ISIN",
    "Rating Full", "Quantity", "Market (Rs. cr)", "YTM", "Maturity Date",
    "Issuer", "Industry", "Sector", "Coupon", "EXTRA 1", "EXTRA 2",
    "EXTRA 3", "Key",
]

# ==============================================================================
# HELPERS
# ==============================================================================

def remove_brackets(series):
    """Remove everything inside () including the brackets themselves."""
    return series.astype(str).str.replace(r'\s*\(.*?\)', '', regex=True).str.strip()

def trim_fund(series):
    """Strip whitespace from fund names, preserving NaN as NaN (never 'nan' string)."""
    return series.where(series.isna(), series.astype(str).str.strip())

def ytm_as_decimal(series):
    """Value is already a decimal (e.g. 0.0727) — keep as-is."""
    return pd.to_numeric(series, errors='coerce')

def ytm_div100(series):
    """Value is a percentage number (e.g. 7.27) — divide by 100."""
    return pd.to_numeric(series, errors='coerce') / 100

def make_output(n):
    return pd.DataFrame({col: [None] * n for col in OUTPUT_COLUMNS})


# ==============================================================================
# BANK 1 — Aditya Birla  (Compiled_AdityaBirla.xlsx)
# B=1 instrument | C=2 ISIN | D=3 rating | E=4 qty | F=5 market
# H=7 YTM (decimal) | J=9 fund name
# ==============================================================================
def load_aditya_birla(path):
    df = pd.read_excel(path, header=None)
    n = len(df)
    out = make_output(n)
    out["Fund Name"]              = trim_fund(df.iloc[:, 9])
    out["Name of the Instrument"] = df.iloc[:, 1]
    out["ISIN"]                   = df.iloc[:, 2]
    out["Rating Full"]            = df.iloc[:, 3]
    out["Quantity"]               = df.iloc[:, 4]
    out["Market (Rs. cr)"]        = df.iloc[:, 5]
    out["YTM"]                    = ytm_as_decimal(df.iloc[:, 7])
    print(f"  ✓ Aditya Birla : {n:,} rows")
    return out


# ==============================================================================
# BANK 2 — Axis Bank  (Compiled_AxisBank.xlsx)
# A=0 fund | B=1 instrument | C=2 ISIN | D=3 rating | E=4 qty
# F=5 market | H=7 YTM (decimal)
# ==============================================================================
def load_axis(path):
    df = pd.read_excel(path, header=None)
    n = len(df)
    out = make_output(n)
    out["Fund Name"]              = trim_fund(df.iloc[:, 0])
    out["Name of the Instrument"] = df.iloc[:, 1]
    out["ISIN"]                   = df.iloc[:, 2]
    out["Rating Full"]            = df.iloc[:, 3]
    out["Quantity"]               = df.iloc[:, 4]
    out["Market (Rs. cr)"]        = df.iloc[:, 5]
    out["YTM"]                    = ytm_as_decimal(df.iloc[:, 7])
    print(f"  ✓ Axis         : {n:,} rows")
    return out


# ==============================================================================
# BANK 3 — Bandhan Bank  (Compiled_BandhanBank.xlsx)
# A=0 fund — three types of bad values to fix, all replaced by nearest fund
#             name found by looking upward:
#             (a) genuinely blank / NaN  (e.g. "Reverse Repo" rows)
#             (b) TRP_*_VAL pattern rows
# B=1 instrument | C=2 ISIN | D=3 rating | E=4 qty | F=5 market
# H=7 YTM (decimal)
# ==============================================================================
def load_bandhan(path):
    df = pd.read_excel(path, header=None)
    fund_col = df.iloc[:, 0].copy()

    # TRP pattern now matches both old format (TRP_160226_VAL) and new (TRP_160426)
    trp_pattern = re.compile(r'^TRP_', re.IGNORECASE)

    def is_bad(val):
        if pd.isna(val):
            return True
        s = str(val).strip()
        return s == "" or trp_pattern.match(s) is not None

    # Walk every row: if bad, replace with the nearest good value above it
    for i in range(len(fund_col)):
        if is_bad(fund_col.iloc[i]):
            for j in range(i - 1, -1, -1):
                if not is_bad(fund_col.iloc[j]):
                    fund_col.iloc[i] = fund_col.iloc[j]
                    break

    n = len(df)
    out = make_output(n)
    out["Fund Name"]              = trim_fund(fund_col)
    out["Name of the Instrument"] = df.iloc[:, 1]
    out["ISIN"]                   = df.iloc[:, 2]
    out["Rating Full"]            = df.iloc[:, 3]
    out["Quantity"]               = df.iloc[:, 4]
    out["Market (Rs. cr)"]        = df.iloc[:, 5]
    out["YTM"]                    = ytm_as_decimal(df.iloc[:, 7])
    print(f"  ✓ Bandhan      : {n:,} rows")
    return out


# ==============================================================================
# BANK 4 — DSP  (Compiled_DSP.xlsx)
# B=1 fund | C=2 instrument | D=3 ISIN | E=4 rating | F=5 qty
# G=6 market | K=10 YTM (divide by 100)
# ==============================================================================
def load_dsp(path):
    df = pd.read_excel(path, header=None)
    n = len(df)
    out = make_output(n)
    out["Fund Name"]              = trim_fund(df.iloc[:, 1])
    out["Name of the Instrument"] = df.iloc[:, 2]
    out["ISIN"]                   = df.iloc[:, 3]
    out["Rating Full"]            = df.iloc[:, 4]
    out["Quantity"]               = df.iloc[:, 5]
    out["Market (Rs. cr)"]        = df.iloc[:, 6]
    out["YTM"]                    = ytm_div100(df.iloc[:, 10])
    print(f"  ✓ DSP          : {n:,} rows")
    return out


# ==============================================================================
# BANK 5 — HDFC  (Compiled_HDFC.xlsx)
# A=0 fund (remove brackets) | B=1 ISIN | D=3 instrument | E=4 rating
# F=5 qty | G=6 market | I=8 YTM (divide by 100)
# ==============================================================================
def load_hdfc(path):
    df = pd.read_excel(path, header=None)
    n = len(df)
    out = make_output(n)
    out["Fund Name"]              = trim_fund(remove_brackets(df.iloc[:, 0]))
    out["ISIN"]                   = df.iloc[:, 1]
    out["Name of the Instrument"] = df.iloc[:, 3]
    out["Rating Full"]            = df.iloc[:, 4]
    out["Quantity"]               = df.iloc[:, 5]
    out["Market (Rs. cr)"]        = df.iloc[:, 6]
    out["YTM"]                    = ytm_div100(df.iloc[:, 8])
    print(f"  ✓ HDFC         : {n:,} rows")
    return out


# ==============================================================================
# BANK 6 — ICICI  (Compiled_ICICI.xlsx)
# A=0 fund | B=1 instrument | C=2 ISIN | E=4 rating | F=5 qty
# G=6 market | I=8 YTM (divide by 100)
# ==============================================================================
def load_icici(path):
    df = pd.read_excel(path, header=None)
    n = len(df)
    out = make_output(n)
    out["Fund Name"]              = trim_fund(df.iloc[:, 0])
    out["Name of the Instrument"] = df.iloc[:, 1]
    out["ISIN"]                   = df.iloc[:, 2]
    out["Rating Full"]            = df.iloc[:, 4]
    out["Quantity"]               = df.iloc[:, 5]
    out["Market (Rs. cr)"]        = df.iloc[:, 6]
    out["YTM"]                    = ytm_div100(df.iloc[:, 8])
    print(f"  ✓ ICICI        : {n:,} rows")
    return out


# ==============================================================================
# BANK 7 — Kotak  (Compiled_Kotak.xlsx)
# Only rows BEFORE the first row where col A == "Scheme".
# J=9  fund name (strip "Portfolio of " and " as on DD-Mon-YYYY" for any date)
# A=0  some instrument names land here incorrectly — copy each non-empty
#      A value into the same row of C, then use the full col C as instrument.
#      Col A is not output anywhere.
# C=2 Name of the instrument | D=3 ISIN | E=4 rating
# F=5 YTM (div by 100) | G=6 qty | H=7 market
# ==============================================================================
def load_kotak(path):
    df = pd.read_excel(path, header=None)

    # Truncate at first row where col A == "Scheme"
    scheme_mask = df.iloc[:, 0].astype(str).str.strip().str.lower() == "scheme"
    if scheme_mask.any():
        df = df.iloc[:scheme_mask.idxmax()].copy()

    # Where col A has a value, copy it into col C of the same row
    col_a = df.iloc[:, 0]
    a_has_value = col_a.notna() & (col_a.astype(str).str.strip() != '')
    df.iloc[a_has_value.values, 2] = col_a[a_has_value].values

    # Fund name: strip "Portfolio of " and any " as on DD-Mon-YYYY" date suffix
    # Using regex for the date so it works regardless of which fortnight's data is used
    fund_raw = df.iloc[:, 9].astype(str)
    fund_clean = (fund_raw
                  .str.replace("Portfolio of ", "", regex=False)
                  .str.replace(r"\s+as on \d{1,2}-[A-Za-z]{3}-\d{4}", "", regex=True))

    n = len(df)
    out = make_output(n)
    out["Fund Name"]              = trim_fund(fund_clean)
    out["Name of the Instrument"] = df.iloc[:, 2]
    out["ISIN"]                   = df.iloc[:, 3]
    out["Rating Full"]            = df.iloc[:, 4]
    out["YTM"]                    = ytm_div100(df.iloc[:, 5])
    out["Quantity"]               = df.iloc[:, 6]
    out["Market (Rs. cr)"]        = df.iloc[:, 7]
    print(f"  ✓ Kotak        : {n:,} rows")
    return out


# ==============================================================================
# BANK 8 — Mirae  (Compiled_Mirae.xlsx)
# A=0 fund — blank cells back-filled from the row immediately BELOW (bfill).
#   Step 1: replace empty strings with NaN so bfill sees them correctly.
#   Step 2: bfill fills each NaN from the value in the row below.
# B=1 instrument | C=2 ISIN | D=3 rating | E=4 qty | F=5 market
# H=7 YTM (decimal)
# ==============================================================================
def load_mirae(path):
    df = pd.read_excel(path, header=None)
    fund_col = df.iloc[:, 0].copy()
    fund_col = fund_col.replace(r'^\s*$', pd.NA, regex=True)  # empty strings → NaN
    fund_col = fund_col.bfill()                                # fill from row below
    n = len(df)
    out = make_output(n)
    out["Fund Name"]              = trim_fund(fund_col)
    out["Name of the Instrument"] = df.iloc[:, 1]
    out["ISIN"]                   = df.iloc[:, 2]
    out["Rating Full"]            = df.iloc[:, 3]
    out["Quantity"]               = df.iloc[:, 4]
    out["Market (Rs. cr)"]        = df.iloc[:, 5]
    out["YTM"]                    = ytm_as_decimal(df.iloc[:, 7])
    print(f"  ✓ Mirae        : {n:,} rows")
    return out


# ==============================================================================
# BANK 9 — Nippon  (Compiled_Nippon.xlsx)
# A=0 fund (remove brackets) | B=1 ISIN | C=2 instrument | D=3 rating
# E=4 qty | F=5 market | H=7 YTM (decimal)
# ==============================================================================
def load_nippon(path):
    df = pd.read_excel(path, header=None)
    n = len(df)
    out = make_output(n)
    out["Fund Name"]              = trim_fund(remove_brackets(df.iloc[:, 0]))
    out["ISIN"]                   = df.iloc[:, 1]
    out["Name of the Instrument"] = df.iloc[:, 2]
    out["Rating Full"]            = df.iloc[:, 3]
    out["Quantity"]               = df.iloc[:, 4]
    out["Market (Rs. cr)"]        = df.iloc[:, 5]
    out["YTM"]                    = ytm_as_decimal(df.iloc[:, 7])
    print(f"  ✓ Nippon       : {n:,} rows")
    return out


# ==============================================================================
# BANK 10 — SBI  (Compiled_SBI.xlsx)
# A=0 fund (remove brackets) | C=2 instrument | D=3 ISIN | E=4 rating
# F=5 qty | G=6 market | I=8 YTM (divide by 100)
# ==============================================================================
def load_sbi(path):
    df = pd.read_excel(path, header=None)
    n = len(df)
    out = make_output(n)
    out["Fund Name"]              = trim_fund(remove_brackets(df.iloc[:, 0]))
    out["Name of the Instrument"] = df.iloc[:, 2]
    out["ISIN"]                   = df.iloc[:, 3]
    out["Rating Full"]            = df.iloc[:, 4]
    out["Quantity"]               = df.iloc[:, 5]
    out["Market (Rs. cr)"]        = df.iloc[:, 6]
    out["YTM"]                    = ytm_div100(df.iloc[:, 8])
    print(f"  ✓ SBI          : {n:,} rows")
    return out


# ==============================================================================
# FILE REGISTRY
# Place all 10 source files in the same folder as this script, then run it.
# ==============================================================================
FILES = [
    ("Compiled_AdityaBirla.xlsx", load_aditya_birla),
    ("Compiled_AxisBank.xlsx",    load_axis),
    ("Compiled_BandhanBank.xlsx", load_bandhan),
    ("Compiled_DSP.xlsx",         load_dsp),
    ("Compiled_HDFC.xlsx",        load_hdfc),
    ("Compiled_ICICI.xlsx",       load_icici),
    ("Compiled_Kotak.xlsx",       load_kotak),
    ("Compiled_Mirae.xlsx",       load_mirae),
    ("Compiled_Nippon.xlsx",      load_nippon),
    ("Compiled_SBI.xlsx",         load_sbi),
]


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    frames = []
    for filename, loader in FILES:
        if not os.path.exists(filename):
            print(f"  ⚠ Skipping (not found): {filename}")
            continue
        try:
            df = loader(filename)
            frames.append(df)
        except Exception as e:
            print(f"  ✗ ERROR in {filename}: {e}")

    if not frames:
        print("No data loaded. Ensure source files are in the same folder as this script.")
        return

    combined = pd.concat(frames, ignore_index=True)[OUTPUT_COLUMNS]

    # --- Market Value: sanitise → numeric → divide by 100 ---
    # Step 1: strip to clean text (removes accidental spaces, comma-separators)
    mkt = combined["Market (Rs. cr)"].astype(str).str.strip().str.replace(',', '', regex=False)
    # Step 2: coerce to numeric — anything non-numeric becomes NaN, then 0
    mkt = pd.to_numeric(mkt, errors='coerce').fillna(0)
    # Step 3: divide by 100
    combined["Market (Rs. cr)"] = mkt / 100

    output_path = "compiled_funds.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        combined.to_excel(writer, index=False, sheet_name="Sheet1")

        ws = writer.sheets["Sheet1"]

        # Format YTM column as percentage
        ytm_col_idx = OUTPUT_COLUMNS.index("YTM") + 1  # 1-based
        for row in ws.iter_rows(min_row=2, min_col=ytm_col_idx, max_col=ytm_col_idx):
            for cell in row:
                if cell.value is not None:
                    cell.number_format = '0.00%'

        # Auto-fit column widths
        for col in ws.columns:
            max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)

    print(f"\n✅ Done! {len(combined):,} total rows written to '{output_path}'")


if __name__ == "__main__":
    main()