# Gurupadigam Weekly Report Generator (v2 — single slot per row)

Automates the weekly Gurupadigam mentee report: one Excel row + one student photo in →
one 2-slide PowerPoint report out (student detail slide + a personalized parent note),
generated in bulk for the whole class.

> **This is v2 of the tool.** It replaces the earlier 4-slots-per-student format with a
> simpler one-slot-per-row format — if a student has multiple slots, add one row per slot.

## How it works

```
Excel (student data)  ──┐
                         ├──►  VBA Macro  ──►  PowerPoint (2 slides per row: details + parent note)
Photos (RegNo.jpg)    ──┘
```

## Repo structure

```
├── macro/
│   └── Gurupadigam_PPT_Generator.bas     ← the exported macro (import into Excel)
├── template/
│   └── Gurupadigam_StudentData_Template.xlsx   ← blank data-entry sheet + Instructions tab
├── photos/
│   └── README.md                         ← photo naming rules
└── output/                               ← generated .pptx files land here (git-ignored)
```

## Columns in the data sheet (in this exact order)

| Col | Field |
|---|---|
| A | Registration Number |
| B | Name |
| C | Personal Interest |
| D | Course Code |
| E | Course Name |
| F | Attendance Percent |
| G | Slot |
| H | Total Marks (max) |
| I | Marks Obtained |
| J | Class Average |

**One row = one slot.** If a student attends 4 slots, give them 4 rows (repeat the
Registration Number, Name, and Personal Interest each time).

**Do not reorder or rename columns** — the macro reads by fixed column position.

## One-time setup (for whoever runs the macro)

1. Open `template/Gurupadigam_StudentData_Template.xlsx`, save a working copy as **.xlsm**
   (macro-enabled workbook), keep the sheet named exactly `Sheet1`.
2. In Excel: **Developer tab → Visual Basic → File → Import File** → select
   `macro/Gurupadigam_PPT_Generator.bas`.
3. In the imported code, update this line to point at your local photos folder:
   ```vb
   photoPath = "F:\gurupadigam_package\photos\"
   ```
4. Enable **Trust access to the VBA project object model** (File → Options → Trust Center →
   Trust Center Settings → Macro Settings) if macros are blocked.
5. Make sure **Microsoft PowerPoint** is installed — the macro drives PowerPoint via
   automation, so this only runs on Windows with Office installed (not Mac, not Excel Online).

## Weekly workflow (for colleagues collecting data)

1. Copy the template, fill in one row per student per slot (see the **Instructions** tab).
2. Collect each student's photo into `/photos`, named as their Registration Number.
3. Open the `.xlsm` workbook, press `Alt+F8`, run **`CreatePPTReports`**.
4. A PowerPoint opens with 2 slides per row (details table + photo, and a personalized
   parent note). It auto-saves as `Gurupadigam_Output.pptx` next to the workbook.

## Notes / known limits

- Windows + Excel + PowerPoint desktop workflow only (VBA + COM automation) — will not run
  on Mac Excel, Excel Online, or Google Sheets.
- The parent-note paragraph is randomly varied each run using `PickOne()`, so bulk reports
  don't read identically.
- Attendance <80% and internal marks <60% trigger a different tone in the parent note
  automatically.
If this is a brand new repo, create an empty one first (no README/license), then upload
these files the same way.
