"""
split.py  –  EPC Automation | Reliance Retail EPC Dept
--------------------------------------------------------
Reads the Master Excel file (Main Backup sheet) and generates
one .xlsx file per PM containing only their sites and columns.

HOW TO RUN:
  Double-click  1_RUN_SPLIT.bat
  OR run:  python split.py
"""

import pandas as pd
import json
import os
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Config ────────────────────────────────────────────────────────────────────
with open("config.json") as f:
    cfg = json.load(f)

MASTER_FILE   = cfg["master_file"]
MASTER_SHEET  = cfg["master_sheet"]
UNIQUE_KEY    = cfg["unique_key"]
OUTPUT_FOLDER = cfg["output_folder"]
PM_LIST       = cfg["pms"]
ID_COLS       = cfg["identifier_columns"]
EDIT_COLS     = cfg["editable_columns"]

# ── Header ────────────────────────────────────────────────────────────────────
print()
print("=" * 60)
print("  EPC AUTOMATION  —  SPLIT")
print("  Reliance Retail | EPC Department")
print("=" * 60)
print()

# ── Load master ───────────────────────────────────────────────────────────────
print(f"  Reading: {MASTER_FILE}  (sheet: '{MASTER_SHEET}') ...")

if not os.path.exists(MASTER_FILE):
    print(f"\n  ERROR: '{MASTER_FILE}' not found in this folder.")
    print("  Make sure the master file is in the same folder as split.py")
    print("  and that its name matches exactly what is in config.json")
    input("\n  Press Enter to close...")
    raise SystemExit

df = pd.read_excel(MASTER_FILE, sheet_name=MASTER_SHEET, dtype=str)
df.columns  = df.columns.str.strip()
df["PM"]    = df["PM"].str.strip()
df[UNIQUE_KEY] = df[UNIQUE_KEY].str.strip()
print(f"  Loaded {len(df):,} rows  |  {len(df.columns)} columns")

# ── Validate columns ──────────────────────────────────────────────────────────
all_needed = ID_COLS + EDIT_COLS
missing    = [c for c in all_needed if c not in df.columns]
if missing:
    print("\n  WARNING — these columns were NOT found in the sheet:")
    for m in missing:
        print(f"    ✗  {m}")
    print("  They will be skipped. Check config.json for typos.")
    print()
    ID_COLS   = [c for c in ID_COLS   if c in df.columns]
    EDIT_COLS = [c for c in EDIT_COLS if c in df.columns]

export_cols = ID_COLS + EDIT_COLS

# ── Create output folder ──────────────────────────────────────────────────────
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
datestamp = datetime.today().strftime("%d-%m-%Y")

# ── Styles ────────────────────────────────────────────────────────────────────
HDR_ID_FILL   = PatternFill("solid", fgColor="1F3864")   # dark navy  — identifier cols
HDR_EDIT_FILL = PatternFill("solid", fgColor="375623")   # dark green — editable cols
EDIT_ROW_FILL = PatternFill("solid", fgColor="E2EFDA")   # light green — data rows (editable)
HDR_FONT      = Font(bold=True, color="FFFFFF", size=10)
BODY_FONT     = Font(size=10)
CENTER        = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT          = Alignment(horizontal="left",   vertical="center", wrap_text=False)
THIN          = Side(style="thin", color="D0D0D0")
THIN_BORDER   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def style_sheet(ws, id_count, edit_count):
    """Apply formatting to a PM worksheet."""
    total_cols = id_count + edit_count

    for col_idx in range(1, total_cols + 1):
        is_edit = col_idx > id_count
        col_letter = get_column_letter(col_idx)

        # Header row
        hdr_cell       = ws.cell(row=1, column=col_idx)
        hdr_cell.fill  = HDR_EDIT_FILL if is_edit else HDR_ID_FILL
        hdr_cell.font  = HDR_FONT
        hdr_cell.alignment = CENTER
        hdr_cell.border    = THIN_BORDER

        # Data rows
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font      = BODY_FONT
            cell.alignment = LEFT
            cell.border    = THIN_BORDER
            if is_edit:
                cell.fill = EDIT_ROW_FILL

        # Column width
        max_len = max(
            (len(str(ws.cell(row=r, column=col_idx).value or ""))
             for r in range(1, min(ws.max_row + 1, 50))),
            default=10
        )
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 32)

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 36

# ── Split ─────────────────────────────────────────────────────────────────────
print()
print(f"  {'PM':<28}  {'Sites':>6}  {'File'}")
print(f"  {'─'*28}  {'─'*6}  {'─'*40}")

results = []
for pm in PM_LIST:
    pm_df = df[df["PM"] == pm][export_cols].reset_index(drop=True)

    if len(pm_df) == 0:
        print(f"  {'⚠  ' + pm:<28}  {'—':>6}  No rows found (check spelling in config.json)")
        results.append((pm, 0, None))
        continue

    safe    = pm.replace(" ", "_").replace("/", "-")
    outpath = os.path.join(OUTPUT_FOLDER, f"{safe}_{datestamp}.xlsx")

    with pd.ExcelWriter(outpath, engine="openpyxl") as writer:
        pm_df.to_excel(writer, index=False, sheet_name="My Sites")
        ws = writer.sheets["My Sites"]
        style_sheet(ws, len(ID_COLS), len(EDIT_COLS))

        # ── README sheet ──────────────────────────────────────────────────────
        rws = writer.book.create_sheet("READ ME")
        lines = [
            ["EPC AUTOMATION — INSTRUCTIONS"],
            [""],
            [f"PM Name       : {pm}"],
            [f"Total sites   : {len(pm_df)}"],
            [f"Generated on  : {datetime.today().strftime('%d %b %Y')}"],
            [""],
            ["WHAT TO FILL:"],
            ["  Go to the 'My Sites' tab."],
            ["  Fill in the GREEN columns only — those are yours to update."],
            ["  The BLUE (dark) columns are read-only reference info."],
            ["  Do NOT change the APEX ID column — it is used to merge back."],
            [""],
            ["WHEN DONE:"],
            ["  Save the file (Ctrl+S)."],
            ["  Put it back in the PM_Files folder on the shared drive."],
            ["  Inform your PM Head / admin that your update is ready."],
            [""],
            ["GREEN COLUMNS YOU NEED TO FILL:"],
        ]
        for ec in EDIT_COLS:
            lines.append([f"  •  {ec}"])
        for row_data in lines:
            rws.append(row_data)
        rws.column_dimensions["A"].width = 58
        rws.sheet_view.showGridLines = False

    print(f"  {'✓  ' + pm:<28}  {len(pm_df):>6}  {outpath}")
    results.append((pm, len(pm_df), outpath))

# ── Summary ───────────────────────────────────────────────────────────────────
ok    = [r for r in results if r[1] > 0]
total = sum(r[1] for r in ok)

print()
print("─" * 60)
print(f"  ✓  {len(ok)} files created  |  {total:,} sites distributed")
print(f"  Folder: {OUTPUT_FOLDER}\\")
print("─" * 60)
print()
input("  Press Enter to close...")
