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
    "Amount":                    "#,##0.########",
    "Network Fee":               "#,##0.########",
    "NetworkFee":                "#,##0.########",
    "Service Fee":               "#,##0.########",
    "Service fee":               "#,##0.########",
    "Inflow / (Outflow) Amount": "#,##0.########",
    "Total Balance":             "#,##0.########",
    "SeqNum":                    "#,##0",
    "Account ID":                "#,##0",
    "ParsedDate":                "DD-MMM-YYYY",
    "ParsedDateUnrounded":       "DD-MMM-YYYY HH:MM:SS",
}

# Applied to any numeric column not listed above (standard comma style)
DEFAULT_NUMERIC_FORMAT = "#,##0.##"

# Header row background colour — light steel blue-grey (RGB hex, BGR byte order for COM)
HEADER_BG_COLOR = 0xD6DCE4

# Maps the Fireblocks 'Asset' field (network identifier) to a display network name.
# Only USDT variants carry a network label; everything else shows blank.
NETWORK_MAP = {
    "TRX_USDT_S2UZ": "TRX",
    "USDT_ERC20":     "ETH",
}


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


def read_accepted_currencies(wb):
    """
    Read the accepted Asset Symbol values from _Lists sheet, Column C (C2 down).
    Returns a set of strings, or an empty set if the column is blank or the
    _Lists sheet doesn't exist yet (first run before the user populates it).
    """
    LISTS_SHEET = "_Lists"
    if LISTS_SHEET not in [s.name for s in wb.sheets]:
        return set()

    ws = wb.sheets[LISTS_SHEET]

    # If C2 is empty there are no accepted currencies defined — skip filtering
    if ws.range("C2").value is None:
        return set()

    # expand("down") selects C2 and every consecutive non-empty cell below it.
    # .value returns a scalar for a single cell or a list for multiple cells.
    values = ws.range("C2").expand("down").value
    if not isinstance(values, list):
        values = [values]

    return {str(v).strip() for v in values if v is not None and str(v).strip()}


def setup_vault_ledger_dropdown(wb, vault_names):
    """
    Keeps the Account Name dropdown in 'Vault Ledger'!C3 up to date.

    Steps:
      1. Write the unique, sorted vault names to a hidden helper sheet (_Lists).
      2. Create a workbook-level named range (VaultNamesList) over those cells.
      3. Apply a data-validation list to 'Vault Ledger'!C3 that reads from
         the named range.

    A named range is used as the source rather than a direct cross-sheet
    address (e.g. _Lists!$A$2:$A$124) because Excel's data-validation
    Formula1 parameter does not reliably accept cross-sheet references
    without a named range wrapper.

    Running this on every refresh means the dropdown automatically gains
    any new vault accounts that appear in the Vault Accounts Report CSV.

    Only column A is cleared and rewritten; columns D/E/F and the user-managed
    AcceptedCurrencies in column C are left untouched.
    """
    vault_names_sorted = sorted(vault_names)
    n = len(vault_names_sorted)

    # ── Hidden helper sheet (_Lists) ──────────────────────────────────────────
    # Holds the deduplicated vault name list used as the dropdown source.
    # Hidden so it doesn't appear in the workbook tab bar.
    LISTS_SHEET = "_Lists"
    sheet_names = [s.name for s in wb.sheets]
    if LISTS_SHEET in sheet_names:
        ws_lists = wb.sheets[LISTS_SHEET]
        ws_lists.range("A:A").clear()       # wipe column A only; D/E/F are untouched
    else:
        ws_lists = wb.sheets.add(LISTS_SHEET)

    # Write vault names as a vertical list: A1 = header, A2..A(n+1) = names.
    # Each inner list is one row, so [[name], [name], ...] writes vertically.
    ws_lists.range("A1").value = "VaultNames"
    ws_lists.range("A1").api.Font.Bold = True
    ws_lists.range((2, 1), (n + 1, 1)).value = [[name] for name in vault_names_sorted]

    # ── Named range (VaultNamesList) ──────────────────────────────────────────
    # Delete any pre-existing definition with this name, then re-add it
    # pointing at the freshly written cells on _Lists.
    RANGE_NAME = "VaultNamesList"
    existing_names = [nm.name for nm in wb.names]
    if RANGE_NAME in existing_names:
        wb.names[RANGE_NAME].delete()
    wb.names.add(RANGE_NAME, f"=_Lists!$A$2:$A${n + 1}")

    # ── Data validation on Vault Ledger!C3 ───────────────────────────────────
    ws_vl = wb.sheets["Vault Ledger"]
    cell  = ws_vl.range("C3")
    cell.api.Validation.Delete()   # remove any previous validation rule first
    cell.api.Validation.Add(
        Type=3,        # xlValidateList  — dropdown populated from a list
        AlertStyle=1,  # xlValidAlertStop — rejects values not in the list
        Formula1="=VaultNamesList",
    )


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



def apply_autofilter(ws):
    """Enable AutoFilter on the header row (row 1) of a worksheet."""
    try:
        if ws.api.AutoFilterMode:
            ws.api.AutoFilterMode = False
    except Exception:
        pass
    # Field=1 is required by the COM API; passing no criteria means show all rows.
    ws.range("A1").api.AutoFilter(Field=1)


def update_vault_ledger(wb, ledger_columns):
    """
    Update the column headers in row 5 of 'Vault Ledger' to match LedgerData.
    Rows 6 onward are left untouched — formulas are managed manually in Excel.
    """
    ws = wb.sheets["Vault Ledger"]
    n_cols = len(ledger_columns)
    header_row = 5

    # Clear and rewrite only the header row
    ws.range((header_row, 1), (header_row, n_cols)).clear()
    ws.range((header_row, 1)).value = ledger_columns
    hdr = ws.range((header_row, 1), (header_row, n_cols))
    hdr.api.Font.Bold      = True
    hdr.api.Interior.Color = HEADER_BG_COLOR


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
for col in ["Amount", "Network Fee", "Service Fee", "ParsedDate", "ParsedDateUnrounded"]:
    if col in tx_df.columns:
        tx_df[col] = pd.to_numeric(tx_df[col], errors="coerce")

# Drop transactions that never completed — these should not appear in the
# Data tab or flow through to LedgerData.
EXCLUDED_STATUSES = {"BLOCKED", "FAILED", "CANCELLED", "REJECTED"}
before = len(tx_df)
tx_df = tx_df[~tx_df["Status"].str.upper().isin(EXCLUDED_STATUSES)].reset_index(drop=True)
excluded = before - len(tx_df)
if excluded:
    print(f"      Excluded {excluded:,} rows with non-completed status")

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
    amount = row.get("Amount")
    common = {
        "TxHash":                      row.get("TxHash"),
        "ParsedDateUnrounded":         row.get("ParsedDateUnrounded"),
        "ParsedDate":                  row.get("ParsedDate"),
        "Asset":                       row.get("Asset"),             # raw identifier e.g. 'TRX_USDT_S2UZ'
        "Asset Symbol":                row.get("Asset Symbol"),      # display symbol e.g. 'USDT', 'TRX'
        "Network":                     NETWORK_MAP.get(str(row.get("Asset") or ""), ""),
        "Direction":                   None,   # filled per side below
        "Note":                        row.get("Note"),
        "Amount":                      amount,
        "Source":                      src_name,
        "Source Type":                 src_type,
        "Source Wallet Address":       row.get("Source Address"),
        "Destination":                 dst_name,
        "Destination Type":            dst_type,
        "Destination Wallet Address":  row.get("Destination Address"),
        "NetworkFee":                  row.get("Network Fee"),
        "Service fee":                 row.get("Service Fee"),
    }

    # Outflow — source vault is sending funds; record it from that vault's perspective
    if src_name in internal_vaults:
        outflow_amount = -amount if pd.notna(amount) else None
        output_rows.append({
            **common,
            "VaultName":             src_name,
            "Direction":             "Outflow",
            "Inflow / (Outflow) Amount": outflow_amount,
        })

    # Inflow — destination vault is receiving funds; record it from that vault's perspective
    if dst_name in internal_vaults:
        output_rows.append({
            **common,
            "VaultName":             dst_name,
            "Direction":             "Inflow",
            "Inflow / (Outflow) Amount": amount,
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
        "Key", "SeqNum", "VaultName", "TxHash", "ParsedDateUnrounded", "ParsedDate",
        "Asset", "Asset Symbol", "Network", "Direction", "Note", "Amount",
        "Source", "Source Type", "Source Wallet Address",
        "Destination", "Destination Type", "Destination Wallet Address",
        "NetworkFee", "Service fee", "Inflow / (Outflow) Amount",
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

    # ── Read accepted currencies from _Lists!C before the sheet is cleared ────
    # This must happen here — setup_vault_ledger_dropdown clears _Lists later.
    accepted_currencies = read_accepted_currencies(wb)
    if accepted_currencies:
        print(f"      Accepted currencies: {', '.join(sorted(accepted_currencies))}")

        # Filter the transaction data to only include accepted asset symbols.
        # This applies to both the Data tab and (via ledger_df) LedgerData.
        before = len(tx_df)
        tx_df = tx_df[tx_df["Asset Symbol"].isin(accepted_currencies)].reset_index(drop=True)
        excluded_curr = before - len(tx_df)
        if excluded_curr:
            print(f"      Excluded {excluded_curr:,} rows with unaccepted asset symbols")

        # Apply the same filter to LedgerData on the Asset Symbol column.
        # Recalculate SeqNum and Key so the sequence stays gap-free after removal.
        ledger_df = ledger_df[ledger_df["Asset Symbol"].isin(accepted_currencies)].copy()
        ledger_df["SeqNum"] = ledger_df.groupby("VaultName").cumcount() + 1
        ledger_df["Key"]    = ledger_df["VaultName"] + "|" + ledger_df["SeqNum"].astype(str)
        ledger_df = ledger_df.reset_index(drop=True)
        print(f"      LedgerData after currency filter: {len(ledger_df):,} rows")

    # Write all three sheets in one loop.
    # ws.clear()                  — removes existing content and formatting
    # ws.range("A1").value = data — bulk write: header + all rows at once (fast)
    # format_sheet(ws, df)        — bold header, number formats
    # apply_autofilter(ws)        — enable column-header filter dropdowns
    for sheet_name, df, label, with_filter in [
        ("Data",       tx_df,     "Data",       True),
        ("VaultData",  vault_df,  "VaultData",  False),
        ("LedgerData", ledger_df, "LedgerData", True),
    ]:
        print(f"      Writing {label} tab ({len(df):,} rows)...", end="", flush=True)
        ws = wb.sheets[sheet_name]
        ws.clear()
        ws.range("A1").value = df_to_values(df)
        format_sheet(ws, df)
        if with_filter:
            apply_autofilter(ws)
        print(" done")

    # ── LedgerData column U: "Spam?" header + per-row formula ────────────────
    # ws.clear() above wipes column U, so the header and formulas are rewritten here.
    if not ledger_df.empty:
        print("      Writing Spam? column in LedgerData...", end="", flush=True)
        ws_ledger = wb.sheets["LedgerData"]
        ws_ledger.range("V1").value = "Spam?"
        ws_ledger.range("V1").api.Font.Bold = True
        ws_ledger.range("V1").api.Interior.Color = HEADER_BG_COLOR
        n_ledger = len(ledger_df)
        ws_ledger.range(f"V2:V{n_ledger + 1}").formula = (
            '=IF(J2="Inflow",IF(L2>XLOOKUP(H2,_Lists!C:C,_Lists!E:E),"","Spam"),"")'
        )
        print(" done")

    print("      Updating Vault Ledger dropdown...", end="", flush=True)
    setup_vault_ledger_dropdown(wb, internal_vaults)
    print(" done")

    print("      Updating Vault Ledger tab...", end="", flush=True)
    update_vault_ledger(wb, list(ledger_df.columns))
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
