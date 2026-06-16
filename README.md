# EPC Automation — Setup & Usage Guide
## Reliance Retail | EPC Department

---

## FOLDER STRUCTURE

Place all these files in ONE folder:

```
EPC_Automation\
│
├── Quarterwise_Target_Master.xlsx   ← Your master file (rename to this, or update config.json)
├── config.json                      ← Settings (PMs, columns, etc.)
├── split.py                         ← Split script
├── merge.py                         ← Merge script
├── 1_RUN_SPLIT.bat                  ← Double-click to split
├── 2_RUN_MERGE.bat                  ← Double-click to merge
├── macro_buttons.vba                ← Optional: paste into Excel for buttons
│
└── PM_Files\                        ← Auto-created. PM files go here.
```

After running merge, you'll also see:
- `Updated_Quarterwise_Target_Master.xlsx` — the merged result
- `Merge_Log_<timestamp>.xlsx` — only created if there were conflicts or warnings

---

## ONE-TIME SETUP

### Step 1 — Install Python
Download from https://python.org/downloads
During install, tick "Add Python to PATH"

(If downloads are blocked on a company PC, try `winget install Python.Python.3.12`
in CMD, or use the portable/embeddable Python zip.)

### Step 2 — Install required libraries
Open Command Prompt and run:
```
pip install pandas openpyxl
```

### Step 3 — Place your master file
Rename your master Excel file to:
```
Quarterwise_Target_Master.xlsx
```
(Or edit `"master_file"` in config.json to match your actual filename)

### Step 4 — Add Excel Buttons (optional)
1. Open your master file
2. Save it as .xlsm (File → Save As → Excel Macro-Enabled Workbook)
3. Press Alt + F11 to open the VBA editor
4. Insert → Module → paste contents of macro_buttons.vba
5. Update the folder path inside the macro
6. Add buttons via Developer tab → Insert → Button (Form Control)

---

## WEEKLY WORKFLOW

### Admin — Start of update cycle:
1. Double-click `1_RUN_SPLIT.bat`
2. 9 files are created in the `PM_Files` folder
3. Share each PM's file with them

### PMs — Their job:
1. Open their file
2. Go to the "My Sites" tab
3. Update any cells as needed
4. Save and return the file to `PM_Files`

### Admin — After all PMs are done:
1. Double-click `2_RUN_MERGE.bat`
2. Watch the terminal — if it shows a CRITICAL WARNING (duplicate APEX ID,
   renamed/added column), read it carefully before typing `CONTINUE`
3. Open `Updated_Quarterwise_Target_Master.xlsx`
   - YELLOW APEX ID = row was updated
   - RED APEX ID = conflict (two PMs touched the same cell — check Merge_Log)

---

## WHAT THE MERGE SCRIPT PROTECTS AGAINST

| Issue | Behaviour |
|---|---|
| PM deletes a cell | Detected and cleared in master, row highlighted |
| Hidden characters / extra spaces | Cleaned automatically before comparing |
| Different date formats (12-06-2026 vs 12/06/2026) | Treated as the same date |
| Case differences (Done vs done) | Treated as the same value |
| Unknown APEX ID (not in master) | Logged as a warning, safely skipped |
| Duplicate APEX ID in one PM file | Admin must type CONTINUE to proceed |
| PM renames/adds a column | Admin must type CONTINUE to proceed |
| Two PMs edit the same cell differently | Flagged as conflict, first value kept, RED highlight |

---

## CONFIG.JSON — HOW TO CHANGE SETTINGS

To add/remove a PM — edit the `"pms"` list (name must match exactly as it
appears in the PM column of the master file).

To change the master filename — edit `"master_file"`.

No other file needs to be touched for these changes.

---

## TROUBLESHOOTING

| Problem | Fix |
|---|---|
| "python not found" | Reinstall Python, tick "Add to PATH" |
| "ModuleNotFoundError: openpyxl" | Run: pip install pandas openpyxl |
| PM_Files not found during merge | Run split.py and merge.py from the SAME folder |
| CMD window closes instantly | Open CMD manually in the folder and run `python merge.py` to see the error |
| Pivot table crash on load | Already fixed — merge.py writes a fresh output file instead |
| File won't save / permission error | Close the file in Excel before running merge |

---

*Built by EPC Automation Team — Reliance Retail Internship Project*
