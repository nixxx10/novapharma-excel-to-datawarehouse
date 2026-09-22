Attribute VB_Name = "modValidate"
Option Explicit
' ============================================================================
' ValidateData - evaluates the Data_Quality sheet, logs every failing check
' and shows a summary. Nothing is changed in the data.
' ============================================================================
Public Sub ValidateData()
    Dim ws As Worksheet, r As Long, lastR As Long, fails As Long, msg As String
    Set ws = ThisWorkbook.Worksheets(SHEET_DQ)
    Application.Calculate
    lastR = LastRow(ws, 2)
    For r = 5 To lastR
        If ws.Cells(r, 5).Value = "CHECK" Then
            fails = fails + 1
            msg = msg & vbCrLf & " - " & ws.Cells(r, 2).Value & ": " & Format(ws.Cells(r, 3).Value, "#,##0")
        End If
    Next r
    If fails = 0 Then
        WriteLog "ValidateData", "PASS: all checks OK"
        MsgBox "Data quality: PASS. All checks OK.", vbInformation
    Else
        WriteLog "ValidateData", "FAIL: " & fails & " check(s) - see Data_Quality"
        ws.Activate
        MsgBox "Data quality: FAIL (" & fails & " check(s))" & msg & vbCrLf & vbCrLf & "Fix Sales_Raw and run RefreshConsolidation again.", vbExclamation
    End If
End Sub

' Optional helper: normalises SKU codes in Sales_Raw (TRIM + UPPER) - the kind of
' "quick fix" that hides data-quality problems instead of solving them at source.
Public Sub NormaliseSkuCodes()
    Dim ws As Worksheet, r As Long, lastR As Long, n As Long, v As String
    Set ws = ThisWorkbook.Worksheets(SHEET_RAW)
    lastR = LastRow(ws)
    SpeedOn
    For r = 2 To lastR
        v = CStr(ws.Cells(r, 3).Value)
        If v <> UCase$(Trim$(v)) Then
            ws.Cells(r, 3).Value = UCase$(Trim$(v))
            n = n + 1
        End If
    Next r
    SpeedOff
    WriteLog "NormaliseSkuCodes", n & " SKU code(s) normalised (TRIM/UPPER)"
    MsgBox n & " SKU code(s) normalised. Run RefreshConsolidation.", vbInformation
End Sub
