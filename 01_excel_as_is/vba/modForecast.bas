Attribute VB_Name = "modForecast"
Option Explicit
' ============================================================================
' SnapshotForecastVersion - creates a new rolling-forecast version:
'   copies the latest version's rows, replaces the months up to the cut-off
'   with actuals (SUMIFS on Sales_Enriched by SKU x Territory x Month),
'   appends the rows to Forecast, re-points Fc_* names and updates Parameters.
' ============================================================================
Public Sub SnapshotForecastVersion()
    Dim wsF As Worksheet, newVer As String, cutoff As Variant, curVer As String
    Set wsF = ThisWorkbook.Worksheets(SHEET_FC)
    curVer = ThisWorkbook.Names("LatestForecast").RefersToRange.Value

    newVer = InputBox("Name of the new forecast version (e.g. FC 2026-Q3):", "Snapshot forecast", "FC 2026-Q3")
    If Len(newVer) = 0 Then Exit Sub
    cutoff = InputBox("Last month with actuals to freeze (YYYYMM):", "Snapshot forecast", ThisWorkbook.Names("CurrentMonth").RefersToRange.Value)
    If Not IsNumeric(cutoff) Then Exit Sub
    cutoff = CLng(cutoff)

    If Application.WorksheetFunction.CountIf(wsF.Columns(1), newVer) > 0 Then
        MsgBox "Version '" & newVer & "' already exists.", vbExclamation: Exit Sub
    End If

    On Error GoTo Fail
    SpeedOn
    Dim lastR As Long, r As Long, dest As Long, n As Long
    lastR = LastRow(wsF)
    dest = lastR + 1
    Dim enr As Worksheet: Set enr = ThisWorkbook.Worksheets(SHEET_ENR)
    Dim rgUnits As Range, rgNet As Range, rgSku As Range, rgTerr As Range, rgMk As Range
    Set rgUnits = ThisWorkbook.Names("Enr_Units").RefersToRange
    Set rgNet = ThisWorkbook.Names("Enr_NetSales").RefersToRange
    Set rgSku = ThisWorkbook.Names("Enr_SKU").RefersToRange
    Set rgTerr = ThisWorkbook.Names("Enr_Territory").RefersToRange
    Set rgMk = ThisWorkbook.Names("Enr_MonthKey").RefersToRange

    For r = 2 To lastR
        If wsF.Cells(r, 1).Value = curVer Then
            wsF.Cells(dest, 1).Value = newVer
            wsF.Cells(dest, 2).Value = wsF.Cells(r, 2).Value
            wsF.Cells(dest, 3).Value = wsF.Cells(r, 3).Value
            wsF.Cells(dest, 4).Value = wsF.Cells(r, 4).Value
            If wsF.Cells(r, 2).Value <= cutoff Then
                wsF.Cells(dest, 5).Value = Application.WorksheetFunction.SumIfs(rgUnits, rgSku, wsF.Cells(r, 3).Value, rgTerr, wsF.Cells(r, 4).Value, rgMk, wsF.Cells(r, 2).Value)
                wsF.Cells(dest, 6).Value = Application.WorksheetFunction.SumIfs(rgNet, rgSku, wsF.Cells(r, 3).Value, rgTerr, wsF.Cells(r, 4).Value, rgMk, wsF.Cells(r, 2).Value)
            Else
                wsF.Cells(dest, 5).Value = wsF.Cells(r, 5).Value
                wsF.Cells(dest, 6).Value = wsF.Cells(r, 6).Value
            End If
            wsF.Cells(dest, 7).FormulaR1C1 = wsF.Cells(2, 7).FormulaR1C1   ' Brand lookup (R1C1 keeps relative refs)
            wsF.Cells(dest, 8).FormulaR1C1 = wsF.Cells(2, 8).FormulaR1C1   ' Region lookup
            dest = dest + 1: n = n + 1
        End If
    Next r
    ResizeColumnNames "Fc_", SHEET_FC, dest - 1
    ThisWorkbook.Names("LatestForecast").RefersToRange.Value = newVer
    On Error Resume Next
    ThisWorkbook.Names("LatestForecast").RefersToRange.Validation.Delete
    On Error GoTo Fail
    Application.Calculate
    WriteLog "SnapshotForecastVersion", "Created " & newVer & " (" & Format(n, "#,##0") & " rows, actuals frozen to " & cutoff & ") from " & curVer
    SpeedOff
    MsgBox "Version " & newVer & " created with " & Format(n, "#,##0") & " rows. Parameters!LatestForecast updated.", vbInformation
    Exit Sub
Fail:
    SpeedOff
    WriteLog "SnapshotForecastVersion", "FAILED: " & Err.Description
    MsgBox "Snapshot failed: " & Err.Description, vbCritical
End Sub
