Attribute VB_Name = "modRefresh"
Option Explicit
' ============================================================================
' RefreshConsolidation - rebuilds Sales_Enriched from Sales_Raw:
'   copies the raw values, fills the lookup/derived formulas down,
'   re-points the Enr_* named ranges, recalculates and stamps LastRefresh.
' ============================================================================
Public Sub RefreshConsolidation()
    Dim wsR As Worksheet, wsE As Worksheet
    Set wsR = ThisWorkbook.Worksheets(SHEET_RAW)
    Set wsE = ThisWorkbook.Worksheets(SHEET_ENR)

    On Error GoTo Fail
    SpeedOn
    Application.StatusBar = "Refreshing Sales_Enriched..."

    Dim lastRaw As Long, lastEnr As Long
    lastRaw = LastRow(wsR)
    lastEnr = LastRow(wsE)
    If lastRaw < 2 Then Err.Raise vbObjectError + 10, , "Sales_Raw is empty."

    ' 0) sort raw by date then ID so the duplicate flag (=A2=A1) works
    wsR.Range("A1").Resize(lastRaw, RAW_COLS).Sort Key1:=wsR.Range("B2"), Order1:=xlAscending, Key2:=wsR.Range("A2"), Order2:=xlAscending, Header:=xlYes

    ' 1) clear everything below the header except row 2 formulas (kept as template)
    If lastEnr > 2 Then wsE.Range(wsE.Cells(3, 1), wsE.Cells(lastEnr, ENR_LAST_COL)).ClearContents

    ' 2) copy raw values into A:H
    wsE.Range("A2").Resize(lastRaw - 1, RAW_COLS).Value = wsR.Range("A2").Resize(lastRaw - 1, RAW_COLS).Value
    wsE.Range("B2").Resize(lastRaw - 1, 1).NumberFormat = "dd-mmm-yyyy"

    ' 3) fill formulas I:S down from the template row 2
    If lastRaw > 2 Then
        wsE.Range(wsE.Cells(2, RAW_COLS + 1), wsE.Cells(lastRaw, ENR_LAST_COL)).FillDown
    End If

    ' 4) re-point named ranges Enr_* to the new size
    ResizeColumnNames "Enr_", SHEET_ENR, lastRaw

    ' 5) recalculate and stamp
    Application.StatusBar = "Recalculating..."
    Application.Calculate
    ThisWorkbook.Names("LastRefresh").RefersToRange.Value = Now
    ThisWorkbook.Names("LastRefresh").RefersToRange.NumberFormat = "dd-mmm-yyyy hh:mm"

    WriteLog "RefreshConsolidation", "Sales_Enriched rebuilt: " & Format(lastRaw - 1, "#,##0") & " rows"
    SpeedOff
    MsgBox "Sales_Enriched rebuilt with " & Format(lastRaw - 1, "#,##0") & " rows." & vbCrLf & "Next step: run ValidateData.", vbInformation
    Exit Sub
Fail:
    SpeedOff
    WriteLog "RefreshConsolidation", "FAILED: " & Err.Description
    MsgBox "Refresh failed: " & Err.Description, vbCritical
End Sub
