# JCI Service Ticket Analysis — Weekly Operational Review

A two-part toolkit for transforming a raw **ServiceNow SAPUI5 export** into a
business-ready Excel workbook for weekly ticket review.

| Part | What it is | When to use |
|------|-----------|-------------|
| **Python script** | `scripts/process_tickets.py` | Run from the command line each week — full rebuild from raw file |
| **Excel VBA workbook** | `vba/modJCITracker.bas` + `JCI Weekly Tracker.xlsm` | Self-contained Excel file with Refresh button — no Python required |

---

## Table of Contents

1. [Background & Problem Statement](#1-background--problem-statement)
2. [How It Works — Overview](#2-how-it-works--overview)
3. [Repository Structure](#3-repository-structure)
4. [Column Reference](#4-column-reference)
5. [Business Logic Reference](#5-business-logic-reference)
6. [Part A — Python Script Setup & Usage](#6-part-a--python-script-setup--usage)
7. [Part B — Excel VBA Workbook Setup & Usage](#7-part-b--excel-vba-workbook-setup--usage)
8. [Weekly Workflow (Step-by-Step)](#8-weekly-workflow-step-by-step)
9. [Carry-Forward Logic](#9-carry-forward-logic)
10. [Protection & Data Integrity](#10-protection--data-integrity)
11. [Troubleshooting](#11-troubleshooting)
12. [FAQ](#12-faq)

---

## 1. Background & Problem Statement

The SAP MaxAttention team reviews open service tickets weekly with the customer
(Johnson Controls Inc. / JCI).  The raw export from ServiceNow:

- Contains **all tickets** (open and closed) — 2,000+ rows
- Has **no aging columns** — reviewers cannot tell at a glance how stale a ticket is
- Has **no priority tier** — SAP's built-in Priority field is not the same as the
  business-facing Customer Priority used in review meetings
- Has **no running log** — there is no column to record what action was agreed in
  each meeting; notes get lost between weeks

This toolkit solves all three problems and produces a workbook that:

- Shows **only in-scope tickets** by default (open statuses), sorted by urgency
- Calculates **Aging** and **Last Update** in days automatically
- Derives **Customer Priority** from a combination of Aging, Status, and SAP Priority
- Preserves **Customer Update** and **SAP Update** notes week-over-week so
  meeting history is never lost

---

## 2. How It Works — Overview

```
┌─────────────────────────────────────────────────────────────────────┐
│                         EACH WEEK                                   │
│                                                                     │
│  1. Export from ServiceNow                                          │
│     ServiceNow → Reports → SAPUI5 Export → Save as .xlsx           │
│                        │                                           │
│                        ▼                                           │
│  2a. Python path                                                    │
│      python process_tickets.py --raw "caseList 27.xlsx"            │
│                             --prev "caseList 26 Updated.xlsx"      │
│                        │                                           │
│      OR                                                             │
│                        │                                           │
│  2b. Excel VBA path                                                 │
│      Open JCI Weekly Tracker.xlsm                                  │
│      Update Config B2 → click "Refresh This Week"                  │
│                        │                                           │
│                        ▼                                           │
│  3. Review output workbook in weekly meeting                        │
│     • In-scope tickets sorted to top (Aging DESC + Priority)       │
│     • Type meeting notes in Customer Update / SAP Update columns   │
│     • Save file → notes stored permanently                         │
│                        │                                           │
│                        ▼                                           │
│  4. Next week: repeat from step 1                                   │
│     Notes from this week automatically appear next week            │
└─────────────────────────────────────────────────────────────────────┘
```

---

## 3. Repository Structure

```
Service-ticket-Analysis/
│
├── scripts/
│   └── process_tickets.py      # Python CLI processor (main script)
│
├── vba/
│   └── modJCITracker.bas       # VBA module for the Excel workbook
│
├── docs/
│   └── architecture.md         # Detailed design decisions & data flow
│
├── sample_data/
│   └── (place a sanitised sample raw export here for testing)
│
├── output/
│   └── (generated files go here — gitignored)
│
├── requirements.txt            # Python dependencies
├── .gitignore
└── README.md                   # This file
```

---

## 4. Column Reference

### Original columns (from ServiceNow — never modified)

| # | Column | Description |
|---|--------|-------------|
| 1 | CASE | ServiceNow case/ticket number — unique key |
| 2 | SUBJECT | Short description of the issue |
| 3 | STATUS | Current ticket status |
| 4 | PRIORITY | SAP-assigned priority (Very High / High / Medium / Low) |
| 5 | INSTALLATION | SAP installation ID and name |
| 6 | SYSTEM NUMBER | SAP system number |
| 7 | SYSTEM | SAP system identifier |
| 8 | COMPONENT | SAP module and component path |
| 9 | REPORTER ID | SAP user ID of the reporter |
| 10 | REPORTER | Full name of the reporter |
| 11 | CREATOR ID | SAP user ID of the creator |
| 12 | CREATOR | Full name of the creator |
| 13 | CUSTOMER ID | SAP customer ID |
| 14 | CUSTOMER | Customer company name |
| 15 | CREATED ON (UTC) | Ticket creation timestamp (UTC) |
| 16 | UPDATED ON (UTC) | Last update timestamp (UTC) |
| 17 | AUTO-CONFIRM DATE | Date ticket auto-closes if no response |
| 18 | SUBMITTED ON | Date submitted to SAP |
| 19 | COMPLETED ON | Date ticket was completed |

### New columns (added by this toolkit — yellow headers)

| # | Column | Type | Description |
|---|--------|------|-------------|
| 20 | **Aging** | Integer (days) | `today − CREATED ON (UTC)` — how old the ticket is |
| 21 | **Last Update** | Integer (days) | `today − UPDATED ON (UTC)` — days since last activity |
| 22 | **Customer Priority** | Text | Business-facing priority tier (see logic below) |
| 23 | **Customer Update** | Free text | Running log — what the customer needs to do / has done |
| 24 | **SAP Update** | Free text | Running log — what SAP needs to do / has done |

---

## 5. Business Logic Reference

### Status scope

Only three statuses are considered "in-scope" for the weekly operational review.
These rows appear at the top of the workbook, sorted by urgency.
All other statuses are preserved in the file but appear below.

| Status | In scope? |
|--------|-----------|
| **Sent to SAP** | ✅ Yes |
| **Customer Action** | ✅ Yes |
| **SAP Proposed Solution** | ✅ Yes |
| In Processing by SAP | ❌ No |
| Sent to SAP Partner | ❌ No |
| Confirmed | ❌ No |
| Confirmed Automatically | ❌ No |
| Not Sent to SAP | ❌ No |

### Customer Priority logic

All **three conditions must be true simultaneously** for a priority to be assigned.
If any condition is not met, the cell is left blank.

| Priority | Aging threshold | Status condition | SAP Priority (color tier) |
|----------|----------------|-----------------|--------------------------|
| **VERY HIGH** | > 7 days | In-scope | Very High → Red |
| **HIGH** | > 30 days | In-scope | High → Yellow |
| **MEDIUM HIGH** | > 60 days | In-scope | Medium → Green |
| *(blank)* | Threshold not met | — | Low or threshold not reached |

**SAP Priority → Color tier mapping** (no Red/Yellow/Green field exists in the raw
export — the SAP Priority column is used as a proxy):

```
Very High  →  Red      →  VERY HIGH   (if Aging > 7)
High       →  Yellow   →  HIGH        (if Aging > 30)
Medium     →  Green    →  MEDIUM HIGH (if Aging > 60)
Low        →  (none)   →  blank
```

### Sorting

1. **Aging — descending** (oldest tickets first)
2. **Customer Priority — custom order**: VERY HIGH → HIGH → MEDIUM HIGH → blank

In-scope rows are always sorted to the top.
Out-of-scope rows follow below, unsorted.

### Date calculation

Both Aging and Last Update are calculated as:

```
value = floor( (today_UTC − date_UTC).total_seconds() / 86400 )
```

- Dates are parsed as UTC datetimes (as exported by ServiceNow)
- Result is a whole number of days
- If a date field is blank or unparseable, the calculated field is left blank
  and a warning is printed in the processing summary

---

## 6. Part A — Python Script Setup & Usage

### Prerequisites

- Python 3.8 or later
- pip

### Installation

```bash
# 1. Clone this repository
git clone https://github.com/Misra-Ravi/Service-ticket-Analysis.git
cd Service-ticket-Analysis

# 2. (Optional but recommended) create a virtual environment
python3 -m venv .venv
source .venv/bin/activate        # Mac / Linux
.venv\Scripts\activate           # Windows

# 3. Install dependencies
pip install -r requirements.txt
```

### Basic usage

```bash
# Minimum — current week raw file only (no carry-forward)
python scripts/process_tickets.py --raw "Downloads/caseList 27.xlsx"

# With carry-forward from previous week's updated file
python scripts/process_tickets.py \
    --raw  "Downloads/caseList 27.xlsx" \
    --prev "Downloads/caseList 26 Updated.xlsx"

# Specify a custom output path
python scripts/process_tickets.py \
    --raw  "Downloads/caseList 27.xlsx" \
    --prev "Downloads/caseList 26 Updated.xlsx" \
    --out  "Weekly Reviews/JCI_2026-07-09.xlsx"

# Override today's date (useful for testing or backdating)
python scripts/process_tickets.py \
    --raw  "Downloads/caseList 27.xlsx" \
    --date 2026-07-09
```

### Command-line arguments

| Argument | Required | Description |
|----------|----------|-------------|
| `--raw` | ✅ Yes | Path to current week's raw ServiceNow `.xlsx` export |
| `--prev` | No | Path to previous week's updated file for carry-forward |
| `--out` | No | Output file path (default: `<raw name> Updated.xlsx` in same folder) |
| `--date` | No | Override today's date as `YYYY-MM-DD` (default: system date) |

### Previous file format detection

The script automatically detects which format the previous file is in:

| Format | Key column | SAP Update column | Customer Update column |
|--------|-----------|-------------------|----------------------|
| New (caseList Updated) | `CASE` | `SAP Update` | `Customer Update` |
| Legacy (JCI Incidents) | `OSS Message` | `MAX UPDATE` | `JCI UPDATE` |

If neither format is detected, carry-forward is skipped and a warning is printed.

### Sample output (console)

```
============================================================
  JCI Service Ticket Weekly Processor
============================================================
  Raw file   : Downloads/caseList 27.xlsx
  Prev file  : Downloads/caseList 26 Updated.xlsx
  Output     : Downloads/caseList 27 Updated.xlsx
  Today      : 2026-07-09
============================================================

Loading raw file...
  2001 data rows | 19 columns
Loading previous week file...
  47 records in carry-forward lookup
  Key column used   : CASE
  Update cols found : SAP=SAP Update, Customer=Customer Update
Processing rows...
Writing output file...

============================================================
  PROCESSING SUMMARY
============================================================
  Total raw rows           : 2001
  In-scope (visible, top)  : 47
  Out-of-scope (below)     : 1954
  Matched carry-forward    : 31
  New rows (blank logs)    : 16
  Date issues              : 0

  Status distribution:
    Confirmed Automatically          1062   
    Confirmed                         850   
    Customer Action                    25  ← IN SCOPE
    SAP Proposed Solution              18  ← IN SCOPE
    In Processing by SAP               11  
    Not Sent to SAP                    31  
    Sent to SAP                         4  ← IN SCOPE

  Customer Priority (in-scope rows):
    VERY HIGH       : 1
    HIGH            : 13
    MEDIUM HIGH     : 19
    blank           : 14
============================================================
```

---

## 7. Part B — Excel VBA Workbook Setup & Usage

The Excel workbook (`JCI Weekly Tracker.xlsm`) is a self-contained file that
lives permanently on your desktop. It does not require Python.

### Prerequisites

- Microsoft Excel 2016 or later (Mac or Windows)
- Macros must be enabled

### One-time setup (~2 minutes)

This setup is done **once**. After that, the workbook works forever with no
further configuration.

#### Step 1 — Enable the Developer tab

**Mac:**
1. Excel menu → **Preferences** → **Ribbon & Toolbar**
2. Check **Developer** in the right-hand list
3. Click **Save**

**Windows:**
1. File → **Options** → **Customize Ribbon**
2. Check **Developer** in the right-hand list
3. Click **OK**

#### Step 2 — Open the VBA Editor

Click the **Developer** tab → **Visual Basic**
(or press `Option+F11` on Mac / `Alt+F11` on Windows)

#### Step 3 — Import the module

1. In the VBA Editor menu: **File → Import File...**
2. Navigate to the repository folder → `vba/`
3. Select **`modJCITracker.bas`**
4. Click **Open**

You should see `modJCITracker` appear in the left panel under **Modules**:

```
VBAProject (JCI Weekly Tracker.xlsm)
  └── Microsoft Excel Objects
        ├── ThisWorkbook
        └── Sheet1 (Current Week)
        ...
  └── Modules
        └── modJCITracker   ← should appear here
```

#### Step 4 — Add Workbook events

In the left panel, double-click **ThisWorkbook**.
A blank code window opens. Paste this exactly:

```vba
Private Sub Workbook_Open()
    ' Hide Update Log so it cannot be unhidden via the UI
    On Error Resume Next
    ThisWorkbook.Sheets("Update Log").Visible = xlSheetVeryHidden
    On Error GoTo 0
    ' Lock workbook structure (prevents sheet deletion/renaming)
    ThisWorkbook.Protect Password:="JCI2026", Structure:=True, Windows:=False
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Auto-save update notes every time the file is saved
    Call SaveUpdatesToLog
End Sub
```

#### Step 5 — Add sheet header guard

In the left panel, double-click **Sheet1 (Current Week)** (or whichever entry
represents the Current Week sheet).
Paste this into the code window:

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Guard header row against accidental changes
    If Target.Row = 1 Then
        Application.EnableEvents = False
        Call ValidateHeaders
        Application.EnableEvents = True
    End If
End Sub
```

#### Step 6 — Save and close the VBA Editor

Press `Cmd+S` (Mac) or `Ctrl+S` (Windows). Save as `.xlsm` if prompted.
Close the VBA Editor window.

#### Step 7 — Set the raw file path

Click the **Config** tab at the bottom of the workbook.
In cell **B2**, enter the full path to your ServiceNow export:

```
/Users/YourName/Downloads/caseList 27.xlsx
```

The workbook will read from this path every time you click **Refresh This Week**.

### Sheet layout

```
JCI Weekly Tracker.xlsm
│
├── Current Week      ← Main review sheet (rebuilt each week)
│   Columns A–S : Original ServiceNow data (locked)
│   Column T    : Aging (locked, auto-calculated)
│   Column U    : Last Update (locked, auto-calculated)
│   Column V    : Customer Priority (locked, auto-calculated)
│   Column W    : Customer Update  ← EDITABLE (yellow header)
│   Column X    : SAP Update       ← EDITABLE (yellow header)
│
├── Config            ← Settings sheet
│   B2 : Raw file path  ← UPDATE THIS EACH WEEK
│   B3 : Today override (leave blank to use system date)
│   B4 : Protection password (JCI2026)
│
└── Update Log        ← Hidden permanent store (xlSheetVeryHidden)
    Col A : CASE number
    Col B : Customer Update
    Col C : SAP Update
    Col D : Last modified timestamp
```

---

## 8. Weekly Workflow (Step-by-Step)

### Python script workflow

```
Monday morning (or whenever the weekly review is due)
│
├── 1. Export raw file from ServiceNow
│        ServiceNow → Reports → Open your saved SAPUI5 Export report
│        Export as Excel → save to Downloads
│
├── 2. Run the script
│        python scripts/process_tickets.py \
│            --raw  "Downloads/caseList 27.xlsx" \
│            --prev "Downloads/caseList 26 Updated.xlsx"
│
├── 3. Open the output file
│        Downloads/caseList 27 Updated.xlsx
│        → 47 active tickets visible, sorted by Aging
│        → Previous weeks' notes already filled in
│
├── 4. During the review meeting
│        → Type action items in Customer Update and SAP Update columns
│        → Ctrl+S / Cmd+S to save progress
│
└── 5. Keep the file — it becomes --prev for next week
```

### Excel VBA workbook workflow

```
Monday morning
│
├── 1. Export raw file from ServiceNow (same as above)
│
├── 2. Update Config sheet cell B2
│        Replace the old file path with the new one
│
├── 3. Go to Current Week sheet → click "Refresh This Week"
│        → Your previous notes are auto-saved first
│        → Fresh data loads, Aging/Priority recalculated
│        → Previous notes reappear automatically
│
├── 4. During the review meeting
│        → Type in Customer Update / SAP Update columns
│        → Click "Save Updates" button or Cmd+S
│
└── 5. The workbook is your permanent tracker — keep it on your Desktop
```

---

## 9. Carry-Forward Logic

The **Customer Update** and **SAP Update** columns are a running log.
Notes must never be lost between weeks.

### Python script carry-forward

The script matches records between the current raw file and the previous updated
file using the **CASE** column (e.g. `24392/2026`) as the unique key.

```
Current raw (2001 rows)        Previous updated file
      │                               │
      │   CASE = "24392/2026"  ←───  CASE = "24392/2026"
      │                               │  Customer Update: "07/01 - needs UAT"
      │                               │  SAP Update: "Waiting on customer"
      ▼                               │
  Output row 24392/2026               │
  Customer Update ←──────────────────┘
  SAP Update      ←──────────────────┘
```

**For new cases** (not in previous file): both columns are left blank for manual entry.

**Previous file format auto-detection:**

| Column in raw format | Column in JCI Incidents (legacy) |
|---------------------|----------------------------------|
| `SAP Update` | `MAX UPDATE` |
| `Customer Update` | `JCI UPDATE` |
| `CASE` | `OSS Message` |

### Excel VBA carry-forward

The VBA workbook uses the hidden **Update Log** sheet as its permanent store.

```
User types a note in Customer Update or SAP Update column
                    │
                    ▼
      Click "Save Updates" or Cmd+S
                    │
                    ▼
      SaveUpdatesToLog() runs automatically
      Writes {CASE → Customer Update, SAP Update} to Update Log
                    │
                    ▼
      Next week: click "Refresh This Week"
      RefreshThisWeek() reads Update Log and pulls notes back in
```

The Update Log is set to `xlSheetVeryHidden` — it cannot be unhidden through
Excel's Format → Unhide menu. Only VBA code can access it.

---

## 10. Protection & Data Integrity

### What is protected

| Layer | What it protects | How |
|-------|-----------------|-----|
| Cell lock | All cells except Customer Update and SAP Update | `Protection(locked=True)` on all cells; `locked=False` on cols 23–24 |
| Sheet protection | Prevents typing in locked cells | `ws.protection.sheet = True`, password `JCI2026` |
| Workbook structure | Prevents sheet deletion/renaming | `ThisWorkbook.Protect Structure:=True` |
| Header guard | Auto-restores any changed header | `Worksheet_Change` event |
| Update Log visibility | Hidden from UI | `xlSheetVeryHidden` — cannot be unhidden without VBA |

### What users CAN do

| Action | Allowed |
|--------|---------|
| Type in Customer Update column | ✅ |
| Type in SAP Update column | ✅ |
| Sort rows using column headers | ✅ |
| Filter using the dropdown arrows | ✅ |
| Click Refresh This Week | ✅ |
| Click Save Updates | ✅ |
| Edit any other column | ❌ Blocked |
| Delete or rename columns | ❌ Blocked |
| Delete or rename sheets | ❌ Blocked |
| Unhide Update Log | ❌ Blocked |

### Header violation behavior

If a header is accidentally changed (e.g. via paste), the `Worksheet_Change`
event fires immediately, shows a warning dialog, and restores the original value:

```
⚠ Header Violation — Restored
─────────────────────────────────────────────────────
WARNING: Column 3 header was changed.

Expected : "STATUS"
Found    : "My Status"

The header has been restored automatically.
Only the Customer Update and SAP Update columns may be edited.
```

---

## 11. Troubleshooting

### Python script

| Problem | Likely cause | Fix |
|---------|-------------|-----|
| `ModuleNotFoundError: openpyxl` | openpyxl not installed | `pip install openpyxl` |
| `Raw file not found` | Wrong path or filename | Check `--raw` argument; use full absolute path |
| `Aging` column is blank | CREATED ON (UTC) column missing or blank in raw file | Check the raw export; re-export from ServiceNow |
| `0 matched carry-forward rows` | Previous file is from a different time period | Normal on first run; notes will accumulate from this week |
| `Expected columns not found` | Raw file has different column names | ServiceNow export settings may have changed; check column names |

### Excel VBA workbook

| Problem | Likely cause | Fix |
|---------|-------------|-----|
| Buttons don't respond | VBA not imported yet | Complete Steps 3–5 in the one-time setup section |
| "File not found" on Refresh | Config B2 path is wrong or file moved | Update cell B2 with the correct path |
| "Macros are disabled" warning | Excel security settings | Excel Preferences → Security → set to "Disable with notification", reopen file, click Enable |
| Refresh wipes notes | Clicked Refresh without saving | Always click "Save Updates" first, or save the file first — but note: Refresh calls SaveUpdatesToLog automatically as its first step |
| Update Log visible in sheet tabs | `Workbook_Open` not pasted yet | Complete Step 4 in the one-time setup section |
| Can't type in any cell | Sheet protection active before VBA setup | Unprotect manually: Review → Unprotect Sheet → password `JCI2026` |

### Running on Windows

The Python script and VBA module work identically on Windows.

For the VBA workbook on Windows:
- Use `Alt+F11` instead of `Option+F11`
- File paths use backslashes: `C:\Users\YourName\Downloads\caseList 27.xlsx`
- Everything else is identical

---

## 12. FAQ

**Q: Do I need both the Python script and the Excel workbook?**
No. They are two independent options. Use whichever fits your workflow:
- Python script → if you're comfortable with the command line and want a reproducible, auditable pipeline
- Excel VBA workbook → if you want a click-button tool with no Python required

**Q: The CASE number format changed. Will matching still work?**
The script and VBA both match on the exact string value in the CASE column.
As long as the same case number appears in both the current and previous file,
matching works regardless of format.

**Q: Can I add more columns to the output?**
Yes — edit `EXPECTED_RAW_HEADERS` and `NEW_COLS` in `process_tickets.py`, and
update the `EXPECTED_HEADERS` constant in `modJCITracker.bas` to match.

**Q: Can I change the Customer Priority thresholds (7 / 30 / 60 days)?**
Yes. In `process_tickets.py`, edit the `customer_priority()` function.
In the VBA module, edit the `If CLng(aging) > N` lines in `RefreshThisWeek()`.

**Q: Can I change the protection password?**
Yes. Change `WB_PASSWORD = "JCI2026"` in `modJCITracker.bas` and
`ws.protection.password = "JCI2026"` in `process_tickets.py`.
Use the same password in both places.

**Q: The previous week's file has different column names (MAX UPDATE / JCI UPDATE).**
This is the legacy JCI Incidents format. The Python script detects this automatically.
The VBA workbook uses the Update Log as its store, so it is not affected by the
previous file's column names.

**Q: How do I onboard a new team member?**
1. Share this repository link
2. They clone it and run `pip install -r requirements.txt`
3. For the VBA workbook, they follow the one-time setup in Section 7
4. That's it — no other configuration needed

---

## Contributing

Pull requests are welcome. For significant changes, open an issue first to discuss
the proposed change.

Please do not commit actual ServiceNow export files containing customer data.
Use the `sample_data/` folder only for sanitised/anonymised test data.

---

## License

Internal use — SAP MaxAttention team.
