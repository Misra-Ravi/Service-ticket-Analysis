# Architecture & Design Decisions

## Data flow

```
ServiceNow SAPUI5 Export (.xlsx)
        │
        │  Raw file — 19 columns, 2000+ rows, all statuses
        ▼
process_tickets.py  /  RefreshThisWeek() VBA
        │
        ├── Load raw rows
        ├── Calculate Aging = today − CREATED ON (UTC)
        ├── Calculate Last Update = today − UPDATED ON (UTC)
        ├── Derive Customer Priority (3-condition logic)
        ├── Look up Customer Update / SAP Update from previous file or Update Log
        ├── Separate in-scope vs out-of-scope rows
        ├── Sort in-scope: Aging DESC → Priority order
        └── Write output with formatting and protection
                │
                ▼
        Output .xlsx
        ├── Rows 2–48    : In-scope (47 active tickets), sorted
        └── Rows 49–2002 : Out-of-scope (closed), unsorted
```

## Key decisions

### Why match on CASE not SUBJECT?
CASE (e.g. `24392/2026`) is a ServiceNow-assigned unique identifier.
SUBJECT is free text and can be edited, duplicated, or truncated.

### Why raw (uncompressed) VBA chunks?
The VBA source text in an `.xlsm` file is stored inside a binary OLE
Compound File (`vbaProject.bin`) using a custom compression algorithm.
Python has no writable OLE library. Raw (uncompressed) chunks are a
valid alternative format that Excel accepts and recompiles from source
on first open. The p-code (compiled bytecode) is discarded — Excel
regenerates it automatically.

### Why is the Update Log xlSheetVeryHidden?
`xlSheetVeryHidden` (value = 2) cannot be reversed through the Excel UI
(Format → Hide/Unhide is greyed out). Only VBA code can set this state.
This prevents accidental editing or deletion of the permanent log.

### Why not use xlwings for full automation?
Excel for Mac does not expose the VBA object model via AppleScript.
`wb.api.VBProject` raises `AttributeError: Unknown property VBProject`
on macOS regardless of xlwings version. This is a Microsoft restriction,
not a Python limitation. On Windows, COM automation works fully.

### Why sort in-scope rows to the top rather than hiding out-of-scope?
openpyxl can write `row_dimensions[n].hidden = True`, but when a
`<filterColumn>` tag is also present in the XML, Excel re-evaluates
the filter on open and overrides the hidden flags, resulting in a blank
sheet. Sorting active rows to the top is more reliable and allows the
user to see all data by scrolling down without any filter interaction.

### Why all-three-conditions for Customer Priority?
The original spec defines VERY HIGH, HIGH, and MEDIUM HIGH each as
requiring Aging + Status + Color tier simultaneously. A ticket with
`High` priority but only 25 days of aging does NOT qualify as HIGH —
it has not yet crossed the 30-day threshold. Assigning priority based
on SAP Priority alone would misrepresent urgency.
