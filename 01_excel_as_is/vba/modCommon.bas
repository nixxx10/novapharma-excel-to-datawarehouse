Attribute VB_Name = "modCommon"
Option Explicit
' ============================================================================
' NovaPharma Iberia - Commercial Model - shared helpers
' ============================================================================

Public Const SHEET_RAW As String = "Sales_Raw"
Public Const SHEET_ENR As String = "Sales_Enriched"
Public Const SHEET_STG As String = "Import_Staging"
Public Const SHEET_LOG As String = "Macro_Log"
Public Const SHEET_DQ As String = "Data_Quality"
Public Const SHEET_FC As String = "Forecast"
Public Const RAW_COLS As Long = 8          ' A:H raw columns
Public Const ENR_LAST_COL As Long = 19     ' A:S in Sales_Enriched

Public Function LastRow(ws As Worksheet, Optional col As Long = 1) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Public Sub WriteLog(macroName As String, msg As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    r = LastRow(ws) + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "dd-mmm-yyyy hh:mm"
    ws.Cells(r, 2).Value = macroName
    ws.Cells(r, 3).Value = Environ("USERNAME")
    ws.Cells(r, 4).Value = msg
End Sub

Public Sub SpeedOn()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With
End Sub

Public Sub SpeedOff()
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
End Sub

' Re-point every workbook name that starts with prefix to rows 2..lastRow of its column
Public Sub ResizeColumnNames(prefix As String, sheetName As String, lastRw As Long)
    Dim nm As Name, ref As String, colLetter As String, p As Long
    For Each nm In ThisWorkbook.Names
        If Left$(nm.Name, Len(prefix)) = prefix Then
            ref = nm.RefersTo                            ' e.g. =Sales_Enriched!$J$2:$J$48164
            p = InStr(ref, "!$")
            If p > 0 Then
                colLetter = Mid$(ref, p + 2, InStr(p + 2, ref, "$") - p - 2)
                nm.RefersTo = "=" & sheetName & "!$" & colLetter & "$2:$" & colLetter & "$" & lastRw
            End If
        End If
    Next nm
End Sub
