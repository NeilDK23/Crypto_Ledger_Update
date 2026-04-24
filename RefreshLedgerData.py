"""
RefreshLedgerData.py
--------------------
Reads the latest Fireblocks export CSVs and writes three tabs in the
Lumetrade Crypto Master Excel workbook:

  Data        — raw transaction rows from the Transaction History Report CSV,
                with two extra derived date columns added.
  VaultData   — vault account rows from the Vault Accounts Report CSV.
  LedgerData  — one row per vault-side per transaction (Inflow or Outflow),
                sorted by vault name then date, with a running sequence number.

Uses xlwings (Excel's own engine) instead of openpyxl to:
  - Load large workbooks in seconds rather than minutes
  - Write thousands of rows at once as a bulk array (much faster than cell-by-cell)
  - Apply formatting through the full Excel API without corruption issues

Usage:
  1. Export fresh CSVs from Fireblocks into the two report folders.
  2. Make sure the Excel file is CLOSED in Excel.
  3. Run:  python RefreshLedgerData.py
"""

import os
import glob
import pandas as pd
from datetime import datetime
import xlwings as xw


# ── Configuration ──────────────────────────────────────────────────────────────
# Paths are relative to the folder where this script lives.

EXCEL_FILE             = "Lumetrade Crypto Master.xlsx"
TRANSACTION_REPORT_DIR = "Transaction History Report"
VAULT_REPORT_DIR       = "Vault Accounts Report"

# ── Formatting settings ────────────────────────────────────────────────────────

# Column-name → Excel number format string.
# Applied to data rows in whichever sheet that column appears in.
#
#   '#,##0.########'    comma-separated thousands, up to 8 decimal places
#                       (trailing zeros are hidden, so 297000 → 297,000)
#   '#,##0'             comma-separated with no decimal places (for integers)
#   'DD-MMM-YYYY'       date display, e.g. 22-Apr-2026
#   'DD-MMM-YYYY HH:MM:SS'  date + time display
NUMBER_FORMATS = {
    "Amount":              "#,##0.########",
    "Network Fee":         "#,##0.########",
    "NetworkFee":          "#,##0.########",
    "Total Balance":       "#,##0.########",
    "SeqNum":              "#,##0",
    "Account ID":          "#,##0",
    "ParsedDate":          "DD-MMM-YYYY",
    "ParsedDateUnrounded": "DD-MMM-YYYY HH:MM:SS",
}

# Applied to any numeric column not listed above (standard comma style)
DEFAULT_NUMERIC_FORMAT = "#,##0.##"

# Header row background colour — light steel blue-grey (RGB hex, BGR byte order for COM)
HEADER_BG_COLOR = 0xD6DCE4


# ── Helper functions ───────────────────────────────────────────────────────────

def find_latest_csv(folder):
    """
    Look inside 'folder' for any .csv file and return the path to the
    most-recently-modified one. Raises FileNotFoundError if the folder is empty.
    Lets you drop a new export in without renaming anything.
    """
    files = glob.glob(os.path.join(folder, "*.csv"))
    if not files:
        raise FileNotFoundError(
            f"No CSV files found in '{folder}'. "
            "Please export from Fireblocks and place the file there."
        )
    return max(files, key=os.path.getmtime)


def date_to_excel_serials(date_str):
    """
    Convert a Fireblocks date string ('22 Apr 2026 08:34:43 GMT') into two
    Excel serial numbers so they can be stored and formatted as dates in Excel.

      serial_day  — whole number = just the date, no time component.
                    Used to display a clean date in the ParsedDate column.
      serial_full — decimal number = date + fractional time of day.
                    Used for accurate chronological sorting in LedgerData.

    Excel counts days from 30 December 1899 (= day 0). Time is stored as
    a fraction of a full day, so 12:00 noon = 0.5, 06:00 = 0.25, etc.

    Returns (None, None) if the input is blank or has an unexpected format.
    """
    if not date_str or pd.isna(date_str):
        return None, None
    try:
        # Parse the string — e.g. '22 Apr 2026 08:34:43 GMT'
        dt = datetime.strptime(str(date_str).strip(), "%d %b %Y %H:%M:%S GMT")
        excel_epoch = datetime(1899, 12, 30)   # Excel's day-zero reference
        delta = dt - excel_epoch
        serial_full = delta.days + delta.seconds / 86400  # include time fraction
        serial_day  = float(delta.days)                   # date only, no time
        return serial_day, serial_full
    except ValueError:
        return None, None   # Unexpected format — return blank rather than crash


def df_to_values(df):
    """
    Convert a DataFrame into a list-of-lists for xlwings bulk write.
    The first inner list is the column headers; the rest are data rows.

    Two conversions are applied to each cell value:
      - NaN / NaT  → None      (so Excel writes a blank cell, not 'nan')
      - numpy types → Python   (so Excel's COM layer receives standard types)
    """
    header = list(df.columns)
    rows = []
    for record in df.itertuples(index=False, name=None):
        row = []
        for v in record:
            if v is None:
                row.append(None)
                continue
            try:
                # pd.isna() catches float NaN, NaT, and None
                if pd.isna(v):
                    row.append(None)
                    continue
            except (TypeError, ValueError):
                pass  # Some types (e.g. strings) raise on isna() — keep as-is
            # numpy scalars expose .item() to convert to a plain Python type
            if hasattr(v, "item"):
                v = v.item()
            row.append(v)
        rows.append(row)
    return [header] + rows


def format_sheet(ws, df):
    """
    Apply formatting to a worksheet after data has been written:
      1. Bold the header row and give it a background colour.
      2. Apply column-specific number formats from NUMBER_FORMATS.
      3. Apply the default comma format to any other numeric column.

    Formatting is applied at column-range level (not cell-by-cell) so it
    runs quickly even for sheets with tens of thousands of rows.

    ws : xlwings Sheet object
    df : the DataFrame that was written to this sheet
    """
    n_cols    = len(df.columns)
    n_rows    = len(df)         # data rows, not counting the header
    last_row  = n_rows + 1      # 1-based row index of the last data row

    # ── Bold header (row 1) ───────────────────────────────────────────────────
    header_range = ws.range((1, 1), (1, n_cols))
    header_range.api.Font.Bold      = True
    header_range.api.Interior.Color = HEADER_BG_COLOR

    if n_rows == 0:
        return  # Nothing more to format if the sheet is empty

    # ── Number formats for data rows (rows 2 onward) ──────────────────────────
    for col_idx, col_name in enumerate(df.columns, start=1):

        # Only format columns that pandas knows contain numbers.
        # Text columns are left as-is (they are already stored as text in Excel).
        if not pd.api.types.is_numeric_dtype(df[col_name].dtype):
            continue

        # Select all data cells in this column (everything below the header)
        data_range = ws.range((2, col_idx), (last_row, col_idx))

        if col_name in NUMBER_FORMATS:
            # Use the specific format defined for this column name
            data_range.api.NumberFormat = NUMBER_FORMATS[col_name]
        else:
            # Fall back to the default comma format for any other numeric column
            data_range.api.NumberFormat = DEFAULT_NUMERIC_FORMAT


# ── Step 1: Find CSV exports ───────────────────────────────────────────────────
print("=" * 60)
print("Lumetrade Crypto Master — Refresh Script")
print("=" * 60)

print("\n[1/5] Locating latest CSV exports...")
tx_csv_path    = find_latest_csv(TRANSACTION_REPORT_DIR)
vault_csv_path = find_latest_csv(VAULT_REPORT_DIR)
print(f"      Transaction report : {os.path.basename(tx_csv_path)}")
print(f"      Vault report       : {os.path.basename(vault_csv_path)}")


# ── Step 2: Load and prepare the Transaction History CSV ──────────────────────
print("\n[2/5] Loading transaction history CSV...")

# Read all columns as strings so pandas does not guess types incorrectly.
# We will convert only the columns we know are numeric.
tx_df = pd.read_csv(tx_csv_path, dtype=str)
print(f"      Loaded {len(tx_df):,} rows, {len(tx_df.columns)} columns")

# Parse the 'Date' column into two Excel serial-number columns.
# serial_day  → stored in ParsedDate          (date-only, for display)
# serial_full → stored in ParsedDateUnrounded (date+time, for sorting)
print("      Parsing dates...")
date_pairs = tx_df["Date"].apply(date_to_excel_serials)
tx_df["ParsedDate"]          = date_pairs.apply(lambda p: p[0])
tx_df["ParsedDateUnrounded"] = date_pairs.apply(lambda p: p[1])

# Convert known numeric columns from strings to actual numbers.
# errors='coerce' turns any unparseable value (e.g. blank) into NaN.
for col in ["Amount", "Network Fee", "ParsedDate", "ParsedDateUnrounded"]:
    if col in tx_df.columns:
        tx_df[col] = pd.to_numeric(tx_df[col], errors="coerce")

print(f"      Date range: {tx_df['Date'].iloc[-1]}  to  {tx_df['Date'].iloc[0]}")


# ── Step 3: Load the Vault Accounts CSV ───────────────────────────────────────
print("\n[3/5] Loading vault accounts CSV...")
vault_df = pd.read_csv(vault_csv_path, dtype=str)
print(f"      Loaded {len(vault_df):,} rows")

# Convert known numeric columns so they are stored as numbers in Excel
for col in ["Total Balance", "Account ID"]:
    if col in vault_df.columns:
        vault_df[col] = pd.to_numeric(vault_df[col], errors="coerce")

# Build the set of internal vault names.
# Any transaction Source or Destination matching one of these names
# will generate a LedgerData entry for that vault.
internal_vaults = set(vault_df["Account Name"].dropna().str.strip().unique())
print(f"      {len(internal_vaults)} unique internal vault accounts found")


# ── Step 4: Build LedgerData ───────────────────────────────────────────────────
print("\n[4/5] Building LedgerData...")
output_rows = []

for _, row in tx_df.iterrows():
    # Skip rows with no Fireblocks Transaction ID — these are not real transactions
    tx_id = row.get("Fireblocks TxId")
    if pd.isna(tx_id) or str(tx_id).strip() == "":
        continue

    src_name = str(row.get("Source", "")).strip()
    dst_name = str(row.get("Destination", "")).strip()
    src_type = str(row.get("Source Type", "")).strip()
    dst_type = str(row.get("Destination Type", "")).strip()

    # Fields shared between the Outflow and Inflow records for this transaction
    common = {
        "ParsedDate":          row.get("ParsedDate"),
        "DateText":            row.get("Date"),          # original human-readable date
        "Asset":               row.get("Asset Symbol"),  # e.g. 'USDT', 'TRX'
        "Amount":              row.get("Amount"),
        "NetworkFee":          row.get("Network Fee"),
        "Note":                row.get("Note"),
        "TxHash":              row.get("TxHash"),
        "Source":              src_name,
        "Destination":         dst_name,
        "ParsedDateUnrounded": row.get("ParsedDateUnrounded"),
    }

    # Outflow — source vault is sending funds; record it from that vault's perspective
    if src_name in internal_vaults:
        output_rows.append({
            **common,
            "VaultName":        src_name,
            "Direction":        "Outflow",
            "CounterpartyType": dst_type,   # who received the funds
            "Counterparty":     dst_name,
        })

    # Inflow — destination vault is receiving funds; record it from that vault's perspective
    if dst_name in internal_vaults:
        output_rows.append({
            **common,
            "VaultName":        dst_name,
            "Direction":        "Inflow",
            "CounterpartyType": src_type,   # who sent the funds
            "Counterparty":     src_name,
        })

# Build DataFrame from the collected rows
ledger_df = pd.DataFrame(output_rows)

if ledger_df.empty:
    print("      WARNING: No ledger rows generated.")
    print("      Check that vault names in the transaction CSV match the vault accounts CSV.")
else:
    # Sort by vault name, then precise timestamp so entries are in exact time order
    ledger_df = ledger_df.sort_values(
        ["VaultName", "ParsedDateUnrounded"]
    ).reset_index(drop=True)

    # Assign a sequential number per vault (1, 2, 3 ...).
    # cumcount() starts at 0, so +1 makes it 1-based.
    ledger_df["SeqNum"] = ledger_df.groupby("VaultName").cumcount() + 1

    # Composite unique key: 'VaultName|SeqNum'
    ledger_df["Key"] = ledger_df["VaultName"] + "|" + ledger_df["SeqNum"].astype(str)

    # Arrange columns in the standard LedgerData layout
    ledger_df = ledger_df[[
        "Key", "VaultName", "SeqNum", "ParsedDate", "DateText",
        "Asset", "Direction", "CounterpartyType", "Counterparty",
        "Amount", "NetworkFee", "Note", "TxHash",
        "Source", "Destination", "ParsedDateUnrounded",
    ]]

    print(f"      {len(ledger_df):,} rows across {ledger_df['VaultName'].nunique()} vaults")


# ── Step 5: Write to Excel via xlwings ────────────────────────────────────────
# xlwings opens the file through Excel's own engine, which is much faster than
# openpyxl for large workbooks. Data is written as a single bulk array rather
# than one cell at a time, and formatting goes through the full Excel COM API.
print(f"\n[5/5] Opening workbook with Excel engine...")

full_path = os.path.abspath(EXCEL_FILE)
if not os.path.exists(full_path):
    raise FileNotFoundError(f"Workbook not found: {full_path}")

# A lock file is created by Excel when it has the workbook open.
# Writing to the file while it is open can cause a save conflict.
lock_file = os.path.join(
    os.path.dirname(full_path), f"~${os.path.basename(EXCEL_FILE)}"
)
if os.path.exists(lock_file):
    print("\n  *** WARNING: file appears to be open in Excel. Close it first. ***")
    if input("  Continue anyway? (y/n): ").strip().lower() != "y":
        raise SystemExit("Aborted.")

# visible=False keeps the Excel window hidden while the script runs.
# The try/finally block ensures the Excel process is always cleaned up,
# even if the script crashes partway through.
app = xw.App(visible=False)
try:
    wb = app.books.open(full_path)
    print("      Opened successfully")

    # Write all three sheets in one loop.
    # ws.clear()                  — removes existing content and formatting
    # ws.range("A1").value = data — bulk write: header + all rows at once (fast)
    # format_sheet(ws, df)        — bold header, number formats
    for sheet_name, df, label in [
        ("Data",       tx_df,     "Data"),
        ("VaultData",  vault_df,  "VaultData"),
        ("LedgerData", ledger_df, "LedgerData"),
    ]:
        print(f"      Writing {label} tab ({len(df):,} rows)...", end="", flush=True)
        ws = wb.sheets[sheet_name]
        ws.clear()
        ws.range("A1").value = df_to_values(df)
        format_sheet(ws, df)
        print(" done")

    print("      Saving...", end="", flush=True)
    wb.save()
    wb.close()
    print(" done")

finally:
    # Always quit the hidden Excel instance, even if an error occurred above
    app.quit()

print()
print("=" * 60)
print("Refresh complete!")
print("=" * 60)
print(f"  Data tab      : {len(tx_df):,} rows")
print(f"  VaultData tab : {len(vault_df):,} rows")
print(f"  LedgerData tab: {len(ledger_df):,} rows")
print()
