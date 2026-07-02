Attribute VB_Name = "modJCITracker"
'
' ============================================================
'  JCI Weekly Tracker — VBA Module
'  File: modJCITracker.bas
'
'  IMPORT INSTRUCTIONS (one-time setup, ~2 minutes):
'  -------------------------------------------------
'  1. Open JCI Weekly Tracker.xlsm
'  2. Press Option+F11 (Mac) or Alt+F11 (Windows) to open VBA Editor
'  3. File > Import File > select this file (modJCITracker.bas)
'  4. In left panel, double-click "ThisWorkbook" and paste the
'     code block marked [PASTE INTO THISWORKBOOK] below.
'  5. In left panel, double-click the "Current Week" sheet module
'     and paste the code block marked [PASTE INTO SHEET MODULE].
'  6. Cmd+S / Ctrl+S to save. Close the VBA Editor.
'
'  WEEKLY USAGE (after one-time setup):
'  -------------------------------------
'  1. Export new raw file from ServiceNow (SAPUI5 Export > .xlsx)
'  2. Update "Config" sheet cell B2 with the new file path
'  3. Click "Refresh This Week" button on the Current Week sheet
'  4. Add/edit notes in Customer Update and SAP Update columns
'  5. Click "Save Updates" or Cmd+S — your notes are stored safely
'
'  PROTECTION:
'  -----------
'  Password : JCI2026
'  Only columns "Customer Update" (col 23) and "SAP Update" (col 24)
'  are editable. All other cells, headers, and sheet structure
'  are locked. The header guard auto-restores any accidental changes.
'
' ============================================================
Option Explicit

' ── Constants ────────────────────────────────────────────────
Private Const SHEET_CW      As String = "Current Week"
Private Const SHEET_LOG     As String = "Update Log"
Private Const SHEET_CFG     As String = "Config"
Private Const WB_PASSWORD   As String = "JCI2026"

Private Const EXPECTED_HEADERS As String = _
    "CASE|SUBJECT|STATUS|PRIORITY|INSTALLATION|SYSTEM NUMBER|SYSTEM|COMPONENT|" & _
    "REPORTER ID|REPORTER|CREATOR ID|CREATOR|CUSTOMER ID|CUSTOMER|" & _
    "CREATED ON (UTC)|UPDATED ON (UTC)|AUTO-CONFIRM DATE|SUBMITTED ON|COMPLETED ON|" & _
    "Aging|Last Update|Customer Priority|Customer Update|SAP Update"

Private Const IN_SCOPE_STATUSES As String = _
    "Sent to SAP|Customer Action|SAP Proposed Solution"


' ════════════════════════════════════════════════════════════
'  PUBLIC MACROS (called by buttons and event handlers)
' ════════════════════════════════════════════════════════════

' ── Refresh This Week ────────────────────────────────────────
' Reads the raw ServiceNow file, recalculates all columns,
' carries forward Customer Update and SAP Update from the
' Update Log, sorts, and rebuilds the Current Week sheet.
Public Sub RefreshThisWeek()
    Dim ws      As Worksheet
    Dim wsCfg   As Worksheet
    Dim wsLog   As Worksheet
    Set wsCfg = ThisWorkbook.Sheets(SHEET_CFG)
    Set ws    = ThisWorkbook.Sheets(SHEET_CW)
    Set wsLog = ThisWorkbook.Sheets(SHEET_LOG)

    ' ── Read config ──
    Dim rawPath As String
    rawPath = Trim(CStr(wsCfg.Range("B2").Value))
    If rawPath = "" Then
        MsgBox "No raw file path set in Config sheet (cell B2)." & vbCrLf & _
               "Please enter the full path to the current week's ServiceNow export.", _
               vbCritical, "Config Missing"
        Exit Sub
    End If
    If Len(Dir(rawPath)) = 0 Then
        MsgBox "Raw file not found:" & vbCrLf & rawPath & vbCrLf & vbCrLf & _
               "Update the path in the Config sheet (cell B2).", _
               vbCritical, "File Not Found"
        Exit Sub
    End If

    Dim todayDate As Date
    Dim todayOvr As String
    todayOvr = Trim(CStr(wsCfg.Range("B3").Value))
    If todayOvr <> "" Then
        todayDate = CDate(todayOvr)
    Else
        todayDate = Date
    End If

    ' ── Auto-save current updates before wiping the sheet ──
    Call SaveUpdatesToLog

    ' ── Unprotect and clear data rows ──
    ws.Unprotect WB_PASSWORD
    Application.ScreenUpdating = False
    Application.EnableEvents   = False

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then ws.Rows("2:" & lastRow).Delete Shift:=xlUp

    ' ── Open raw file ──
    Dim wbRaw As Workbook
    Dim wsRaw As Worksheet
    Set wbRaw = Workbooks.Open(rawPath, ReadOnly:=True)
    Set wsRaw = wbRaw.Sheets(1)

    ' ── Map raw column positions ──
    Dim rawLastCol As Integer
    rawLastCol = wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column
    Dim c As Integer
    Dim cCase As Integer, cStatus As Integer, cPriority As Integer
    Dim cCreated As Integer, cUpdated As Integer
    For c = 1 To rawLastCol
        Select Case Trim(CStr(wsRaw.Cells(1, c).Value))
            Case "CASE":             cCase     = c
            Case "STATUS":           cStatus   = c
            Case "PRIORITY":         cPriority = c
            Case "CREATED ON (UTC)": cCreated  = c
            Case "UPDATED ON (UTC)": cUpdated  = c
        End Select
    Next c

    If cCase = 0 Or cStatus = 0 Then
        MsgBox "Could not find required columns (CASE, STATUS) in the raw file." & vbCrLf & _
               "Please verify this is a valid ServiceNow SAPUI5 Export.", _
               vbCritical, "Invalid File Format"
        wbRaw.Close SaveChanges:=False
        ws.Protect Password:=WB_PASSWORD, DrawingObjects:=False, _
            Contents:=True, Scenarios:=False, _
            AllowSorting:=True, AllowFiltering:=True
        Application.ScreenUpdating = True
        Application.EnableEvents   = True
        Exit Sub
    End If

    Dim rawLastRow As Long
    rawLastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row

    ' ── Build carry-forward lookup from Update Log ──
    Dim logLookup As Object
    Set logLookup = CreateObject("Scripting.Dictionary")
    Dim logLast As Long, lr As Long
    logLast = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    For lr = 2 To logLast
        Dim lKey As String
        lKey = Trim(CStr(wsLog.Cells(lr, 1).Value))
        If lKey <> "" Then
            logLookup(lKey) = Array(wsLog.Cells(lr, 2).Value, wsLog.Cells(lr, 3).Value)
        End If
    Next lr

    Dim scopeArr() As String
    scopeArr = Split(IN_SCOPE_STATUSES, "|")

    ' ── Process rows into in-scope and out-of-scope arrays ──
    Dim inData()  As Variant
    Dim outData() As Variant
    Dim inCnt  As Long
    Dim outCnt As Long
    inCnt = 0: outCnt = 0
    ReDim inData(1 To rawLastRow, 1 To 24)
    ReDim outData(1 To rawLastRow, 1 To 24)

    Dim r As Long
    For r = 2 To rawLastRow
        Dim caseNum  As String
        Dim status   As String
        Dim priority As String
        caseNum  = Trim(CStr(wsRaw.Cells(r, cCase).Value))
        status   = Trim(CStr(wsRaw.Cells(r, cStatus).Value))
        priority = Trim(CStr(wsRaw.Cells(r, cPriority).Value))

        ' Aging = today minus CREATED ON (UTC)
        Dim aging As Variant
        Dim createdVal As Variant
        createdVal = wsRaw.Cells(r, cCreated).Value
        If IsDate(createdVal) And createdVal <> "" Then
            aging = CLng(todayDate - CDate(createdVal))
        Else
            aging = ""
        End If

        ' Last Update = today minus UPDATED ON (UTC)
        Dim lastUpd As Variant
        Dim updatedVal As Variant
        updatedVal = wsRaw.Cells(r, cUpdated).Value
        If IsDate(updatedVal) And updatedVal <> "" Then
            lastUpd = CLng(todayDate - CDate(updatedVal))
        Else
            lastUpd = ""
        End If

        ' Customer Priority — all 3 conditions must be true
        Dim cp As String
        cp = ""
        Dim inScope As Boolean
        inScope = False
        Dim s As Integer
        For s = 0 To UBound(scopeArr)
            If status = scopeArr(s) Then inScope = True: Exit For
        Next s
        If inScope And IsNumeric(aging) Then
            Select Case priority
                Case "Very High": If CLng(aging) > 7  Then cp = "VERY HIGH"
                Case "High":      If CLng(aging) > 30 Then cp = "HIGH"
                Case "Medium":    If CLng(aging) > 60 Then cp = "MEDIUM HIGH"
            End Select
        End If

        ' Carry-forward from Update Log
        Dim custUpd As Variant
        Dim sapUpd  As Variant
        custUpd = "": sapUpd = ""
        If logLookup.exists(caseNum) Then
            custUpd = logLookup(caseNum)(0)
            sapUpd  = logLookup(caseNum)(1)
        End If

        ' Build output row (19 original + 5 new = 24 columns)
        Dim outRow(1 To 24) As Variant
        Dim oc As Integer
        For oc = 1 To rawLastCol
            outRow(oc) = wsRaw.Cells(r, oc).Value
        Next oc
        outRow(20) = aging
        outRow(21) = lastUpd
        outRow(22) = cp
        outRow(23) = custUpd
        outRow(24) = sapUpd

        If inScope Then
            inCnt = inCnt + 1
            Dim ic As Integer
            For ic = 1 To 24: inData(inCnt, ic) = outRow(ic): Next ic
        Else
            outCnt = outCnt + 1
            Dim ooc As Integer
            For ooc = 1 To 24: outData(outCnt, ooc) = outRow(ooc): Next ooc
        End If
    Next r

    wbRaw.Close SaveChanges:=False

    ' ── Sort in-scope: Aging DESC, then Customer Priority ──
    Dim i As Long, j As Long
    For i = 1 To inCnt - 1
        For j = i + 1 To inCnt
            Dim agI As Long, agJ As Long, prI As Integer, prJ As Integer
            agI = IIf(inData(i, 20) = "", -1, CLng(inData(i, 20)))
            agJ = IIf(inData(j, 20) = "", -1, CLng(inData(j, 20)))
            prI = PriorityOrder(CStr(inData(i, 22)))
            prJ = PriorityOrder(CStr(inData(j, 22)))
            If agI < agJ Or (agI = agJ And prI > prJ) Then
                Dim k As Integer, tmp As Variant
                For k = 1 To 24
                    tmp = inData(i, k): inData(i, k) = inData(j, k): inData(j, k) = tmp
                Next k
            End If
        Next j
    Next i

    ' ── Write rows: in-scope first, then out-of-scope ──
    Dim writeRow As Long
    writeRow = 2
    For i = 1 To inCnt
        For c = 1 To 24: ws.Cells(writeRow, c).Value = inData(i, c): Next c
        writeRow = writeRow + 1
    Next i
    For i = 1 To outCnt
        For c = 1 To 24: ws.Cells(writeRow, c).Value = outData(i, c): Next c
        writeRow = writeRow + 1
    Next i

    ' ── Reapply formatting ──
    Call ApplyFormatting(ws, writeRow - 1)

    ' ── Reprotect sheet ──
    ws.Protect Password:=WB_PASSWORD, DrawingObjects:=False, _
        Contents:=True, Scenarios:=False, _
        AllowSorting:=True, AllowFiltering:=True

    Application.ScreenUpdating = True
    Application.EnableEvents   = True

    MsgBox "Refresh complete!" & vbCrLf & vbCrLf & _
           "In-scope rows : " & inCnt  & vbCrLf & _
           "Total rows    : " & (inCnt + outCnt) & vbCrLf & _
           "Date          : " & CStr(todayDate), _
           vbInformation, "JCI Tracker — Done"
End Sub


' ── Save Updates to Log ──────────────────────────────────────
' Writes any filled Customer Update / SAP Update values from
' Current Week back to the hidden Update Log sheet for permanent
' storage. Called automatically on Save and before every Refresh.
Public Sub SaveUpdatesToLog()
    Dim ws    As Worksheet
    Dim wsLog As Worksheet
    Set ws    = ThisWorkbook.Sheets(SHEET_CW)
    Set wsLog = ThisWorkbook.Sheets(SHEET_LOG)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' Find column positions dynamically (safe against future reorders)
    Dim cCase As Integer, cCU As Integer, cSU As Integer
    Dim hdr As Range
    For Each hdr In ws.Rows(1).Cells
        Select Case hdr.Value
            Case "CASE":            cCase = hdr.Column
            Case "Customer Update": cCU   = hdr.Column
            Case "SAP Update":      cSU   = hdr.Column
        End Select
        If cCase > 0 And cCU > 0 And cSU > 0 Then Exit For
    Next hdr
    If cCase = 0 Or cCU = 0 Or cSU = 0 Then Exit Sub

    wsLog.Unprotect WB_PASSWORD

    ' Build index of existing log entries
    Dim logLast As Long
    logLast = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    Dim logIdx As Object
    Set logIdx = CreateObject("Scripting.Dictionary")
    Dim lr As Long
    For lr = 2 To logLast
        Dim k As String
        k = Trim(CStr(wsLog.Cells(lr, 1).Value))
        If k <> "" Then logIdx(k) = lr
    Next lr

    ' Write back all non-blank update values
    Dim r As Long
    For r = 2 To lastRow
        Dim caseNum As String
        caseNum = Trim(CStr(ws.Cells(r, cCase).Value))
        Dim cu As Variant, su As Variant
        cu = ws.Cells(r, cCU).Value
        su = ws.Cells(r, cSU).Value
        If caseNum = "" Or (cu = "" And su = "") Then GoTo NextRow
        If logIdx.exists(caseNum) Then
            wsLog.Cells(logIdx(caseNum), 2).Value = cu
            wsLog.Cells(logIdx(caseNum), 3).Value = su
            wsLog.Cells(logIdx(caseNum), 4).Value = Now()
        Else
            Dim newRow As Long
            newRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
            wsLog.Cells(newRow, 1).Value = caseNum
            wsLog.Cells(newRow, 2).Value = cu
            wsLog.Cells(newRow, 3).Value = su
            wsLog.Cells(newRow, 4).Value = Now()
            logIdx(caseNum) = newRow
        End If
        NextRow:
    Next r

    wsLog.Protect Password:=WB_PASSWORD
End Sub


' ── Header Validator ─────────────────────────────────────────
' Called by the Worksheet_Change event on Current Week.
' If any header in row 1 is changed, it is immediately restored
' and the user is warned. Prevents accidental column deletion
' or renaming that would break the Refresh logic.
Public Sub ValidateHeaders()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_CW)
    Dim expectedArr() As String
    expectedArr = Split(EXPECTED_HEADERS, "|")
    Dim i As Integer
    For i = 0 To UBound(expectedArr)
        Dim cellVal As String
        cellVal = Trim(CStr(ws.Cells(1, i + 1).Value))
        If cellVal <> expectedArr(i) Then
            MsgBox "WARNING: Column " & (i + 1) & " header was changed." & vbCrLf & vbCrLf & _
                   "Expected : """ & expectedArr(i) & """" & vbCrLf & _
                   "Found    : """ & cellVal & """" & vbCrLf & vbCrLf & _
                   "The header has been restored automatically." & vbCrLf & _
                   "Only the Customer Update and SAP Update columns may be edited.", _
                   vbCritical, "Header Violation — Restored"
            Application.EnableEvents = False
            ws.Unprotect WB_PASSWORD
            ws.Cells(1, i + 1).Value = expectedArr(i)
            ws.Protect Password:=WB_PASSWORD, DrawingObjects:=False, _
                Contents:=True, Scenarios:=False, _
                AllowSorting:=True, AllowFiltering:=True
            Application.EnableEvents = True
            Exit Sub
        End If
    Next i
End Sub


' ════════════════════════════════════════════════════════════
'  PRIVATE HELPERS
' ════════════════════════════════════════════════════════════

Private Function PriorityOrder(cp As String) As Integer
    Select Case cp
        Case "VERY HIGH":   PriorityOrder = 0
        Case "HIGH":        PriorityOrder = 1
        Case "MEDIUM HIGH": PriorityOrder = 2
        Case Else:          PriorityOrder = 99
    End Select
End Function

Public Sub ApplyFormatting(ws As Worksheet, lastDataRow As Long)
    Dim r As Long, c As Integer
    For r = 2 To lastDataRow
        If r Mod 2 = 0 Then
            ws.Rows(r).Interior.Color = RGB(234, 241, 251)
        Else
            ws.Rows(r).Interior.Color = RGB(255, 255, 255)
        End If
        With ws.Rows(r).Borders
            .LineStyle = xlContinuous
            .Color     = RGB(170, 170, 170)
            .Weight    = xlThin
        End With
        For c = 20 To 24
            ws.Cells(r, c).WrapText         = True
            ws.Cells(r, c).VerticalAlignment = xlTop
        Next c
    Next r
End Sub


' ════════════════════════════════════════════════════════════
'
'  ┌──────────────────────────────────────────────────────┐
'  │  PASTE INTO THISWORKBOOK MODULE                      │
'  └──────────────────────────────────────────────────────┘
'
'  Private Sub Workbook_Open()
'      ' Hide Update Log so it cannot be unhidden via the UI
'      On Error Resume Next
'      ThisWorkbook.Sheets("Update Log").Visible = xlSheetVeryHidden
'      On Error GoTo 0
'      ' Lock workbook structure (prevents sheet deletion/renaming)
'      ThisWorkbook.Protect Password:="JCI2026", Structure:=True, Windows:=False
'  End Sub
'
'  Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'      ' Auto-save update notes every time the file is saved
'      Call SaveUpdatesToLog
'  End Sub
'
'
'  ┌──────────────────────────────────────────────────────┐
'  │  PASTE INTO "Current Week" SHEET MODULE              │
'  └──────────────────────────────────────────────────────┘
'
'  Private Sub Worksheet_Change(ByVal Target As Range)
'      ' Guard header row against accidental changes
'      If Target.Row = 1 Then
'          Application.EnableEvents = False
'          Call ValidateHeaders
'          Application.EnableEvents = True
'      End If
'  End Sub
'
' ════════════════════════════════════════════════════════════
