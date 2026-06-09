"""
ClearLedgerData.py
------------------
Clears all data previously imported by RefreshLedgerData.py from the
Lumetrade Crypto Master workbook, while preserving any Excel formulas
and content in areas not written by the refresh script.

Sheets cleared:
  Data                  — row 2 downwards, all columns
  LedgerData            — row 2 downwards, all columns
  VaultData             — row 2 downwards, columns A:G only (H:N recon formulas preserved)
  USDT Master           — row 2 downwards, all columns
  ETH Master            — row 2 downwards, all columns
  TRX Master            — row 2 downwards, all columns
  USDT Lumetrade Master    — row 2 downwards, all columns
  ETH Lumetrade Master     — row 2 downwards, all columns
  TRX Lumetrade Master     — row 2 downwards, all columns
  USDT Lumetrade Treasury  — row 2 downwards, all columns
  ETH Lumetrade Treasury   — row 2 downwards, all columns
  TRX Lumetrade Treasury   — row 2 downwards, all columns

Row 1 headers and any formulas outside the cleared ranges are untouched.

Usage:
  1. Make sure the Excel file is CLOSED in Excel.
  2. Run:  python ClearLedgerData.py
"""

import os
import xlwings as xw

EXCEL_FILE = "Lumetrade Crypto Master.xlsx"

# Sheets where every column from row 2 downwards is cleared
FULL_CLEAR_SHEETS = [
    "Data",
    "LedgerData",
    "USDT Master",
    "ETH Master",
    "TRX Master",
    "USDT Lumetrade Master",
    "ETH Lumetrade Master",
    "TRX Lumetrade Master",
    "USDT Lumetrade Treasury",
    "ETH Lumetrade Treasury",
    "TRX Lumetrade Treasury",
]


def disable_autofilter(ws):
    """Turn off AutoFilter so hidden rows are included in the clear operation."""
    try:
        if ws.api.AutoFilterMode:
            ws.api.AutoFilterMode = False
    except Exception:
        pass


def clear_data_rows(ws, last_col=None):
    """
    Clear content from row 2 downwards across the used range.
    If last_col is provided (1-based), only columns 1..last_col are cleared.
    clear_contents() removes values and formulas but leaves cell formatting intact.
    """
    disable_autofilter(ws)
    used = ws.used_range
    last_row = used.last_cell.row
    if last_col is None:
        last_col = used.last_cell.column
    if last_row < 2:
        return  # Nothing below the header to clear
    ws.range((2, 1), (last_row, last_col)).clear_contents()


# ── Main ──────────────────────────────────────────────────────────────────────

print("=" * 60)
print("Lumetrade Crypto Master — Clear Script")
print("=" * 60)

full_path = os.path.abspath(EXCEL_FILE)
if not os.path.exists(full_path):
    raise FileNotFoundError(f"Workbook not found: {full_path}")

lock_file = os.path.join(
    os.path.dirname(full_path), f"~${os.path.basename(EXCEL_FILE)}"
)
if os.path.exists(lock_file):
    print("\n  *** WARNING: file appears to be open in Excel. Close it first. ***")
    if input("  Continue anyway? (y/n): ").strip().lower() != "y":
        raise SystemExit("Aborted.")

print(f"\nOpening workbook...")
app = xw.App(visible=False)
try:
    wb = app.books.open(full_path)
    sheet_names = [s.name for s in wb.sheets]
    print("  Opened successfully\n")

    # Clear all columns from row 2 downwards
    for sheet_name in FULL_CLEAR_SHEETS:
        if sheet_name in sheet_names:
            print(f"  Clearing {sheet_name}...", end="", flush=True)
            clear_data_rows(wb.sheets[sheet_name])
            print(" done")
        else:
            print(f"  Skipping '{sheet_name}' (sheet not found)")

    # VaultData: only columns A:G from row 2 — H:N recon formulas are preserved
    if "VaultData" in sheet_names:
        print("  Clearing VaultData (A:G only)...", end="", flush=True)
        clear_data_rows(wb.sheets["VaultData"], last_col=7)
        print(" done")
    else:
        print("  Skipping 'VaultData' (sheet not found)")

    print("\n  Saving...", end="", flush=True)
    wb.save()
    wb.close()
    print(" done")

finally:
    app.quit()

print()
print("=" * 60)
print("Clear complete!")
print("=" * 60)
