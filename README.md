# EPC Automation — Setup & Usage Guide
## Reliance Retail | EPC Department

---

## FOLDER STRUCTURE

Place all these files in ONE folder (e.g. C:\EPC_Automation\):

```
EPC_Automation\
│
├── Quarterwise_Target_Master.xlsx   ← Master file (rename yours to this)
├── config.json                      ← Settings (PMs, columns, etc.)
├── split.py                         ← Split script
├── merge.py                         ← Merge script
├── 1_RUN_SPLIT.bat                  ← Double-click to split
├── 2_RUN_MERGE.bat                  ← Double-click to merge
├── macro_buttons.vba                ← Paste this into Excel (see below)
│
├── PM_Files\                        ← Auto-created. PM files go here.
└── Backups\                         ← Auto-created. Backups + audit logs.
```

---

## ONE-TIME SETUP

### Step 1 — Install Python
Download from https://python.org/downloads
During install, tick "Add Python to PATH"

### Step 2 — Install required libraries
Open Command Prompt and run:
```
pip install pandas openpyxl
```

### Step 3 — Rename your master file
Rename the master Excel file to:
```
Quarterwise_Target_Master.xlsx
```
(Or change "master_file" in config.json to match your filename)

### Step 4 — Add Excel Buttons (optional but recommended)
1. Open your master file
2. Save it as .xlsm  (File → Save As → Excel Macro-Enabled Workbook)
3. Press Alt + F11 to open the VBA editor
4. Click Insert → Module
5. Paste the contents of macro_buttons.vba
6. Change EPC_FOLDER to your actual folder path
7. Close the editor
8. Developer tab → Insert → Button → draw on sheet → assign RunSplit or RunMerge

---

## WEEKLY WORKFLOW

### Admin — Start of update cycle:
1. Double-click `1_RUN_SPLIT.bat`  (or click "Generate PM Files" button in Excel)
2. 6 files are created in the PM_Files folder
3. Share each PM's file with them (email / shared drive)

### PMs — Their job:
1. Open their file (e.g. Jitendra_Talreja_15-06-2025.xlsx)
2. Go to "My Sites" tab
3. Fill in the GREEN columns only
4. Save and return the file to the PM_Files folder

### Admin — After all PMs are done:
1. Double-click `2_RUN_MERGE.bat`  (or click "Merge Updates" button in Excel)
2. Master file is updated automatically
3. Yellow cells = updated, Red cells = conflict (two PMs edited same cell)
4. Check Backups\ folder for audit log if any conflicts

---

## CONFIG.JSON — HOW TO CHANGE SETTINGS

To add a new PM:
- Open config.json in Notepad
- Add their name to the "pms" list (must match exactly as it appears in the PM column)

To change editable columns:
- Edit the "editable_columns" list in config.json

To change the master file name:
- Update "master_file" in config.json

---

## TROUBLESHOOTING

| Problem | Fix |
|---|---|
| "python not found" | Reinstall Python and tick "Add to PATH" |
| "Module not found" | Run: pip install pandas openpyxl |
| Master file not found | Make sure filename matches config.json exactly |
| PM name not matching | Check for extra spaces in the PM column in Excel |
| Merge takes too long | Normal for large files — wait for it to finish |

---

*Built by EPC Automation Team — Reliance Retail Internship Project*
