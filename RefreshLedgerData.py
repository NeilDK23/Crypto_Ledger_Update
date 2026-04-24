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

Usage:
  1. Export fresh CSVs from Fireblocks into the two report folders.
  2. Make sure the Excel file is CLOSED in Excel.
  3. Run:  python RefreshLedgerData.py
"""

import os
import glob
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook


# ── Configuration ──────────────────────────────────────────────────────────────
# These paths are relative to the folder where this script lives.
# Change them here if you move files around.

EXCEL_FILE              = "Lumetrade Crypto Master.xlsx"
TRANSACTION_REPORT_DIR  = "Transaction History Report"
VAULT_REPORT_DIR        = "Vault Accounts Report"


# ── Helper functions ───────────────────────────────────────────────────────────

def find_latest_csv(folder):
    """
    Look inside 'folder' for any .csv file and return the path to the
    most-recently-modified one. Raises an error if the folder is empty.
    This means you can drop a new export in without renaming anything.
    """
    csv_files = glob.glob(os.path.join(folder, "*.csv"))
    if not csv_files:
        raise FileNotFoundError(
            f"No CSV files found in '{folder}'. "
            "Please export from Fireblocks and place the file there."
        )
    # max() with os.path.getmtime as the key picks the newest file
    return max(csv_files, key=os.path.getmtime)


def date_to_excel_serials(date_str):
    """
    Convert a Fireblocks date string such as '22 Apr 2026 08:34:43 GMT'
    into two Excel serial numbers:

      serial_day  — whole-number serial for just the date (time stripped off).
                    Useful for grouping transactions by day.
      serial_full — decimal serial including the time of day as a fraction.
                    Useful for sorting transactions in exact time order.

    Excel counts days from 30 December 1899 (day 0). A time component is
    stored as the decimal fraction of a 24-hour day, e.g. 06:00 = 0.25.

    Returns (None, None) if the string is blank or cannot be parsed.
    """
    # Treat blank / NaN values as missing
    if not date_str or pd.isna(date_str):
        return None, None
    try:
        # Parse the string into a Python datetime object
        dt = datetime.strptime(str(date_str).strip(), "%d %b %Y %H:%M:%S GMT")
        # Excel's reference point — day 0 is 30 Dec 1899
        excel_epoch = datetime(1899, 12, 30)
        delta = dt - excel_epoch
        # Full serial: whole days plus fraction-of-day for the time
        serial_full = delta.days + delta.seconds / 86400
        # Day-only serial: just the whole number (drops the time)
        serial_day = float(delta.days)
        return serial_day, serial_full
    except ValueError:
        # If the format is unexpected, return blanks rather than crashing
        return None, None


def clean_value(v):
    """
    Convert NaN / NaT values to Python None so openpyxl writes a blank
    cell instead of the string 'nan'. Leaves all other values unchanged.
    """
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        # pd.isna() can raise for some object types — treat those as valid
        pass
    return v


def write_dataframe_to_sheet(ws, df):
    """
    Clear the worksheet and write 'df' into it, starting with a header row.
    NaN values are converted to blank cells (via clean_value).
    """
    # Delete every existing row so we start with a clean slate
    ws.delete_rows(1, ws.max_row)

    # Write the column names as the first (header) row
    ws.append(list(df.columns))

    # Write each data row, replacing NaN with None for clean output
    for record in df.itertuples(index=False, name=None):
        ws.append([clean_value(v) for v in record])


# ── Step 1: Locate the CSV files ───────────────────────────────────────────────
print("=" * 60)
print("Lumetrade Crypto Master — Refresh Script")
print("=" * 60)

print("\n[1/6] Locating latest CSV exports...")

# Find the newest CSV in each report folder automatically
tx_csv_path    = find_latest_csv(TRANSACTION_REPORT_DIR)
vault_csv_path = find_latest_csv(VAULT_REPORT_DIR)

print(f"      Transaction report : {os.path.basename(tx_csv_path)}")
print(f"      Vault account report: {os.path.basename(vault_csv_path)}")


# ── Step 2: Load and prepare the Transaction History CSV ──────────────────────
print("\n[2/6] Loading transaction history CSV...")

# Read every column as a string to avoid pandas guessing wrong types.
# We'll convert numeric columns ourselves below.
tx_df = pd.read_csv(tx_csv_path, dtype=str)
print(f"      Loaded {len(tx_df):,} rows, {len(tx_df.columns)} columns")

# ── Parse the 'Date' column into two numeric Excel serial columns ──────────────
# We need two versions:
#   ParsedDate          — day-level serial (used for grouping / display)
#   ParsedDateUnrounded — full-precision serial (used for exact sorting)
print("      Parsing dates...")
date_pairs = tx_df["Date"].apply(date_to_excel_serials)
tx_df["ParsedDate"]          = date_pairs.apply(lambda pair: pair[0])
tx_df["ParsedDateUnrounded"] = date_pairs.apply(lambda pair: pair[1])

# Convert the columns that should be numbers from strings to floats.
# errors="coerce" turns unparseable values (e.g. blank strings) into NaN.
for col in ["Amount", "Network Fee", "ParsedDate", "ParsedDateUnrounded"]:
    if col in tx_df.columns:
        tx_df[col] = pd.to_numeric(tx_df[col], errors="coerce")

print(f"      Date range: {tx_df['Date'].iloc[-1]}  to  {tx_df['Date'].iloc[0]}")


# ── Step 3: Load the Vault Accounts CSV ───────────────────────────────────────
print("\n[3/6] Loading vault accounts CSV...")

# Read the vault CSV — all columns as strings
vault_df = pd.read_csv(vault_csv_path, dtype=str)
print(f"      Loaded {len(vault_df):,} rows, {len(vault_df.columns)} columns")

# Convert 'Total Balance' to a number so it sorts / sums correctly in Excel
if "Total Balance" in vault_df.columns:
    vault_df["Total Balance"] = pd.to_numeric(vault_df["Total Balance"], errors="coerce")

# Build the set of internal vault names from the 'Account Name' column.
# Any transaction whose Source or Destination matches one of these names
# will generate a ledger entry for that vault.
internal_vaults = set(
    vault_df["Account Name"].dropna().str.strip().unique()
)
print(f"      Found {len(internal_vaults)} unique internal vault accounts")


# ── Step 4: Open the Excel workbook ───────────────────────────────────────────
print(f"\n[4/6] Opening workbook: {EXCEL_FILE}")

# If Excel has the file open, there will be a temporary lock file.
# Writing to the file while it is open can corrupt it, so warn the user.
lock_file = f"~${EXCEL_FILE}"
if os.path.exists(lock_file):
    print()
    print("  *** WARNING ***")
    print(f"  The file '{EXCEL_FILE}' appears to be open in Excel.")
    print("  Close it before running this script to avoid data loss.")
    answer = input("  Continue anyway? (y/n): ").strip().lower()
    if answer != "y":
        print("  Aborted.")
        raise SystemExit(1)

# Load the workbook preserving all formulas and external links.
# Do NOT use data_only=True — openpyxl drops the formula calculation chain
# in that mode, which corrupts external link references when the file is saved.
wb = load_workbook(EXCEL_FILE)
print("      Workbook loaded successfully")


# ── Step 5: Update the Data tab ────────────────────────────────────────────────
print("\n[5/6] Updating tabs...")
print("      Writing Data tab...")

# The Data tab stores every raw transaction row from the CSV,
# plus the two calculated ParsedDate columns we added above.
ws_data = wb["Data"]
write_dataframe_to_sheet(ws_data, tx_df)
print(f"      Data tab: {len(tx_df):,} rows written")


# ── Step 6: Update the VaultData tab ──────────────────────────────────────────
print("      Writing VaultData tab...")

# The VaultData tab mirrors the vault accounts report exactly.
ws_vault = wb["VaultData"]
write_dataframe_to_sheet(ws_vault, vault_df)
print(f"      VaultData tab: {len(vault_df):,} rows written")


# ── Step 7: Build the LedgerData rows ─────────────────────────────────────────
print("      Building LedgerData...")

# We create one output row for every vault that is *involved* in a transaction:
#   - If an internal vault is the SOURCE      -> record an Outflow for that vault
#   - If an internal vault is the DESTINATION -> record an Inflow  for that vault
#
# A transfer between two internal vaults will therefore generate TWO rows:
# one Outflow for the sending vault, and one Inflow for the receiving vault.
output_rows = []

for _, row in tx_df.iterrows():
    # Skip rows with no Fireblocks TxId — these are not real transactions
    tx_id = row.get("Fireblocks TxId")
    if pd.isna(tx_id) or str(tx_id).strip() == "":
        continue

    src_name = str(row.get("Source", "")).strip()
    dst_name = str(row.get("Destination", "")).strip()
    src_type = str(row.get("Source Type", "")).strip()
    dst_type = str(row.get("Destination Type", "")).strip()

    # Fields that are the same whether we are recording an inflow or outflow
    common = {
        "ParsedDate":          row.get("ParsedDate"),
        "DateText":            row.get("Date"),           # human-readable date
        "Asset":               row.get("Asset Symbol"),   # e.g. "USDT", "TRX"
        "Amount":              row.get("Amount"),
        "NetworkFee":          row.get("Network Fee"),
        "Note":                row.get("Note"),
        "TxHash":              row.get("TxHash"),
        "Source":              src_name,
        "Destination":         dst_name,
        "ParsedDateUnrounded": row.get("ParsedDateUnrounded"),
    }

    # Outflow: money leaving an internal vault
    if src_name in internal_vaults:
        output_rows.append({
            **common,
            "VaultName":        src_name,
            "Direction":        "Outflow",
            "CounterpartyType": dst_type,   # who received the funds
            "Counterparty":     dst_name,
        })

    # Inflow: money arriving into an internal vault
    if dst_name in internal_vaults:
        output_rows.append({
            **common,
            "VaultName":        dst_name,
            "Direction":        "Inflow",
            "CounterpartyType": src_type,   # who sent the funds
            "Counterparty":     src_name,
        })

# Convert the list of row-dicts into a DataFrame
ledger_df = pd.DataFrame(output_rows)

if ledger_df.empty:
    # This usually means vault names in the transaction CSV don't match
    # the Account Names in the vault accounts CSV — worth investigating.
    print("      WARNING: No ledger rows were generated.")
    print("      Check that vault names in the transaction CSV match the vault accounts CSV.")
else:
    # Sort by vault name first, then by exact date so each vault's entries
    # are in chronological order before we assign sequence numbers
    ledger_df = ledger_df.sort_values(
        ["VaultName", "ParsedDate"]
    ).reset_index(drop=True)

    # Assign a sequential number (1, 2, 3 ...) within each vault's transactions.
    # cumcount() starts at 0, so we add 1 to make it start at 1.
    ledger_df["SeqNum"] = ledger_df.groupby("VaultName").cumcount() + 1

    # Composite key: "VaultName|SeqNum" — uniquely identifies each ledger row
    ledger_df["Key"] = ledger_df["VaultName"] + "|" + ledger_df["SeqNum"].astype(str)

    # Arrange columns in the standard LedgerData layout
    ledger_df = ledger_df[[
        "Key", "VaultName", "SeqNum", "ParsedDate", "DateText",
        "Asset", "Direction", "CounterpartyType", "Counterparty",
        "Amount", "NetworkFee", "Note", "TxHash",
        "Source", "Destination", "ParsedDateUnrounded",
    ]]


# ── Step 8: Write LedgerData tab ──────────────────────────────────────────────
print("      Writing LedgerData tab...")
ws_ledger = wb["LedgerData"]
write_dataframe_to_sheet(ws_ledger, ledger_df)
print(f"      LedgerData tab: {len(ledger_df):,} rows written")


# ── Step 9: Save the workbook ─────────────────────────────────────────────────
print(f"\n[6/6] Saving workbook...")
wb.save(EXCEL_FILE)

print()
print("=" * 60)
print("Refresh complete!")
print("=" * 60)
print(f"  Data tab      : {len(tx_df):,} rows")
print(f"  VaultData tab : {len(vault_df):,} rows")
print(f"  LedgerData tab: {len(ledger_df):,} rows")
print()
