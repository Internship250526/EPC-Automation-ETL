"""
merge.py  –  EPC Automation | Reliance Retail EPC Dept
--------------------------------------------------------
Reads all PM files from PM_Files folder and merges their
editable columns back into the Master file.

HOW TO RUN:
  Double-click  2_RUN_MERGE.bat
  OR run:  python merge.py
"""

import pandas as pd
import json
import os
import shutil
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ── Config ────────────────────────────────────────────────────────────────────
with open("config.json") as f:
    cfg = json.load(f)

MASTER_FILE   = cfg["master_file"]
MASTER_SHEET  = cfg["master_sheet"]
UNIQUE_KEY    = cfg["unique_key"]
OUTPUT_FOLDER = cfg["output_folder"]
EDIT_COLS     = cfg["editable_columns"]

CHANGED_FILL  = PatternFill("solid", fgColor="FFFF00")   # yellow  — value updated
CONFLICT_FILL = PatternFill("solid", fgColor="FFB3B3")   # red     — conflict (two PMs touched same cell)
CHANGED_FONT  = Font(size=10, bold=True)

# ── Header ────────────────────────────────────────────────────────────────────
print()
print("=" * 60)
print("  EPC AUTOMATION  —  MERGE")
print("  Reliance Retail | EPC Department")
print("=" * 60)
print()

# ── Backup master ─────────────────────────────────────────────────────────────
if not os.path.exists(MASTER_FILE):
    print(f"  ERROR: '{MASTER_FILE}' not found.")
    print("  Make sure the master file is in the same folder as merge.py")
    input("\n  Press Enter to close...")
    raise SystemExit

os.makedirs("Backups", exist_ok=True)
stamp       = datetime.today().strftime("%d-%m-%Y_%H%M%S")
backup_path = f"Backups/BACKUP_{stamp}_{MASTER_FILE}"
shutil.copy2(MASTER_FILE, backup_path)
print(f"  Backup saved  →  {backup_path}")

# ── Load master into pandas ───────────────────────────────────────────────────
print(f"  Reading master: {MASTER_FILE} ...")
master_df = pd.read_excel(MASTER_FILE, sheet_name=MASTER_SHEET, dtype=str)
master_df.columns  = master_df.columns.str.strip()
master_df[UNIQUE_KEY] = master_df[UNIQUE_KEY].str.strip()
print(f"  Loaded {len(master_df):,} rows\n")

# ── Collect PM files ──────────────────────────────────────────────────────────
if not os.path.exists(OUTPUT_FOLDER):
    print(f"  ERROR: '{OUTPUT_FOLDER}' folder not found.")
    print("  Run split.py first to generate PM files.")
    input("\n  Press Enter to close...")
    raise SystemExit

pm_files = sorted([
    f for f in os.listdir(OUTPUT_FOLDER)
    if f.endswith(".xlsx") and not f.startswith("~$")
])

if not pm_files:
    print(f"  ERROR: No .xlsx files found in '{OUTPUT_FOLDER}'.")
    input("\n  Press Enter to close...")
    raise SystemExit

print(f"  Found {len(pm_files)} PM file(s) in '{OUTPUT_FOLDER}':\n")

# ── Merge logic ───────────────────────────────────────────────────────────────
# pending_changes[apex_id][col] = {"value": ..., "source": filename}
# If two files update the same cell → conflict
pending  = {}   # apex_id → {col: {value, source}}
conflict_log = []
skipped_ids  = []

for filename in pm_files:
    filepath = os.path.join(OUTPUT_FOLDER, filename)
    print(f"  Processing: {filename}")

    try:
        import openpyxl as _oxl
        _wb = _oxl.load_workbook(filepath, read_only=True)
        sheet_name = "My Sites" if "My Sites" in _wb.sheetnames else _wb.sheetnames[0]
        _wb.close()
        pm_df = pd.read_excel(filepath, sheet_name=sheet_name, dtype=str)
        print(f"    Reading sheet: '{sheet_name}'")
    except Exception as e:
        print(f"    ⚠  Could not read — {e}\n")
        continue

    pm_df.columns  = pm_df.columns.str.strip()

    if UNIQUE_KEY not in pm_df.columns:
        print(f"    ⚠  '{UNIQUE_KEY}' column missing — skipping\n")
        continue

    pm_df[UNIQUE_KEY] = pm_df[UNIQUE_KEY].str.strip()
    cols_present = [c for c in EDIT_COLS if c in pm_df.columns]

    updated = 0
    for _, row in pm_df.iterrows():
        apex_id = row[UNIQUE_KEY]
        if not apex_id or apex_id == "nan":
            continue

        if apex_id not in master_df[UNIQUE_KEY].values:
            skipped_ids.append({"File": filename, UNIQUE_KEY: apex_id,
                                 "Issue": "APEX ID not found in master"})
            continue

        if apex_id not in pending:
            pending[apex_id] = {}

        for col in cols_present:
            new_val = str(row.get(col, "")).strip()
            if new_val in ("", "nan", "NaT"):
                continue  # PM left it blank — don't overwrite master

            if col in pending[apex_id]:
                # Another PM already updated this same cell → conflict
                prev = pending[apex_id][col]
                if prev["value"] != new_val:
                    conflict_log.append({
                        "APEX ID"     : apex_id,
                        "Column"      : col,
                        "Value from"  : prev["source"],
                        "Value"       : prev["value"],
                        "Conflict from": filename,
                        "Conflict value": new_val,
                        "Resolution"  : "Conflict — kept FIRST value, flagged RED in master"
                    })
                    # Keep first value; mark as conflict
                    pending[apex_id][col]["conflict"] = True
                    continue
            else:
                pending[apex_id][col] = {
                    "value"   : new_val,
                    "source"  : filename,
                    "conflict": False
                }
            updated += 1

    print(f"    ✓  {updated} cell updates queued\n")

# ── Apply changes to master workbook (openpyxl — preserves formatting) ────────
print("  Applying changes to master workbook ...")
wb = load_workbook(MASTER_FILE)

if MASTER_SHEET not in wb.sheetnames:
    print(f"  ERROR: Sheet '{MASTER_SHEET}' not in master workbook.")
    input("\n  Press Enter to close...")
    raise SystemExit

ws = wb[MASTER_SHEET]

# Build column index map from header row (row 1)
col_map = {}
for cell in ws[1]:
    if cell.value:
        col_map[str(cell.value).strip()] = cell.column

# Build APEX ID → Excel row map
apex_col = col_map.get(UNIQUE_KEY)
if not apex_col:
    print(f"  ERROR: '{UNIQUE_KEY}' column not found in master sheet.")
    input("\n  Press Enter to close...")
    raise SystemExit

apex_to_row = {}
for row_idx in range(2, ws.max_row + 1):
    v = ws.cell(row=row_idx, column=apex_col).value
    if v:
        apex_to_row[str(v).strip()] = row_idx

# Write changes
cells_changed   = 0
cells_conflict  = 0

for apex_id, col_changes in pending.items():
    excel_row = apex_to_row.get(apex_id)
    if not excel_row:
        continue

    for col, info in col_changes.items():
        col_idx = col_map.get(col)
        if not col_idx:
            continue

        cell       = ws.cell(row=excel_row, column=col_idx)
        cell.value = info["value"]   # always write — no condition

        if info["conflict"]:
            cell.fill = CONFLICT_FILL   # red   — conflict, first value kept
            cells_conflict += 1
        else:
            cell.fill = CHANGED_FILL    # yellow — cleanly updated
            cell.font = CHANGED_FONT

        cells_changed += 1

wb.save(MASTER_FILE)
print(f"  Master saved  →  {MASTER_FILE}")
print(f"  Cells updated: {cells_changed}  |  Conflicts: {cells_conflict}")

# ── Audit log ─────────────────────────────────────────────────────────────────
log_rows = conflict_log + skipped_ids
if log_rows:
    log_path = f"Backups/Audit_Log_{stamp}.xlsx"
    pd.DataFrame(log_rows).to_excel(log_path, index=False)
    print(f"\n  ⚠  Audit log saved  →  {log_path}")
    if conflict_log:
        print(f"     {len(conflict_log)} conflict(s) — highlighted RED in master")
    if skipped_ids:
        print(f"     {len(skipped_ids)} APEX ID(s) not found in master")
else:
    print("\n  ✓  No conflicts. Clean merge!")

# ── Summary ───────────────────────────────────────────────────────────────────
print()
print("─" * 60)
print(f"  MERGE COMPLETE")
print(f"  Cells updated  :  {cells_changed}")
print(f"  Conflicts (RED):  {cells_conflict}")
print(f"  Yellow cells   :  clean updates")
print(f"  RED cells      :  conflict — check Audit_Log in Backups\\")
print("─" * 60)
print()
input("  Press Enter to close...")
