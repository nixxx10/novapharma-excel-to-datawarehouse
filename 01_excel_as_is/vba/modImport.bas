Attribute VB_Name = "modImport"
Option Explicit
' ============================================================================
' ImportMonthlyFile - loads the monthly ERP extract (CSV) into Import_Staging,
' validates the header, appends the rows to Sales_Raw and logs the load.
' ============================================================================
Public Sub ImportMonthlyFile()
    Dim fPath As Variant
    fPath = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Select the ERP sales extract (CSV)")
    If VarType(fPath) = vbBoolean Then Exit Sub

    Dim wsS As Worksheet, wsR As Worksheet, wbCsv As Workbook, wsC As Worksheet
    Set wsS = ThisWorkbook.Worksheets(SHEET_STG)
    Set wsR = ThisWorkbook.Worksheets(SHEET_RAW)

    On Error GoTo Fail
    SpeedOn
    wsS.Range(wsS.Cells(2, 1), wsS.Cells(wsS.Rows.Count, RAW_COLS)).ClearContents

    Set wbCsv = Workbooks.Open(Filename:=fPath, Local:=True, ReadOnly:=True)
    Set wsC = wbCsv.Worksheets(1)
    Dim lastR As Long, i As Long
    lastR = LastRow(wsC)
    If lastR < 2 Then Err.Raise vbObjectError + 1, , "The file has no data rows."

    ' 1) header validation
    For i = 1 To RAW_COLS
        If UCase$(Trim$(CStr(wsC.Cells(1, i).Value))) <> UCase$(CStr(wsS.Cells(1, i).Value)) Then
            Err.Raise vbObjectError + 2, , "Header mismatch in column " & i & ": expected '" & wsS.Cells(1, i).Value & "', found '" & wsC.Cells(1, i).Value & "'."
        End If
    Next i

    ' 2) land in staging (values only)
    wsS.Range("A2").Resize(lastR - 1, RAW_COLS).Value = wsC.Range("A2").Resize(lastR - 1, RAW_COLS).Value
    wbCsv.Close SaveChanges:=False
    Set wbCsv = Nothing

    ' 3) basic checks before appending: no blank IDs, dates are dates, units numeric
    Dim r As Long, bad As Long, minKey As Long, maxKey As Long, k As Long
    minKey = 999999: maxKey = 0
    For r = 2 To lastR
        If Len(Trim$(CStr(wsS.Cells(r, 1).Value))) = 0 Or Not IsDate(wsS.Cells(r, 2).Value) Or Not IsNumeric(wsS.Cells(r, 5).Value) Then
            bad = bad + 1
        Else
            k = Year(wsS.Cells(r, 2).Value) * 100 + Month(wsS.Cells(r, 2).Value)
            If k < minKey Then minKey = k
            If k > maxKey Then maxKey = k
        End If
    Next r
    If bad > 0 Then Err.Raise vbObjectError + 3, , bad & " row(s) have a blank ID, an invalid date or non-numeric units. Nothing was appended."

    ' 4) guard against loading the same period twice
    Dim existing As Double
    existing = Application.WorksheetFunction.CountIfs(wsR.Columns(2), ">=" & CDbl(DateSerial(minKey \ 100, minKey Mod 100, 1)), wsR.Columns(2), "<=" & CDbl(DateSerial(maxKey \ 100, maxKey Mod 100 + 1, 0)))
    If existing > 0 Then
        If MsgBox("Sales_Raw already contains " & Format(existing, "#,##0") & " rows for period " & minKey & "-" & maxKey & "." & vbCrLf & "Append anyway?", vbYesNo + vbExclamation) = vbNo Then
            SpeedOff: Exit Sub
        End If
    End If

    ' 5) append
    Dim dest As Long
    dest = LastRow(wsR) + 1
    wsR.Cells(dest, 1).Resize(lastR - 1, RAW_COLS).Value = wsS.Range("A2").Resize(lastR - 1, RAW_COLS).Value
    wsR.Cells(dest, 2).Resize(lastR - 1, 1).NumberFormat = "dd-mmm-yyyy"

    WriteLog "ImportMonthlyFile", "Appended " & Format(lastR - 1, "#,##0") & " rows for period " & minKey & IIf(maxKey <> minKey, "-" & maxKey, "") & " (file " & Dir(fPath) & ")"
    SpeedOff
    MsgBox Format(lastR - 1, "#,##0") & " rows appended to Sales_Raw." & vbCrLf & "Next step: update Parameters!CurrentMonth and run RefreshConsolidation.", vbInformation
    Exit Sub
Fail:
    If Not wbCsv Is Nothing Then wbCsv.Close SaveChanges:=False
    SpeedOff
    WriteLog "ImportMonthlyFile", "FAILED: " & Err.Description
    MsgBox "Import failed: " & Err.Description, vbCritical
End Sub
