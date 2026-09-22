Attribute VB_Name = "modExport"
Option Explicit
' ============================================================================
' ExportToCSV - writes the master and fact sheets to UTF-8 CSV files
' (values only) in a folder chosen by the user. These are the files the
' Python ETL of the migration project consumes.
' ============================================================================
Public Sub ExportToCSV()
    Dim sheets As Variant, i As Long, folder As String, ws As Worksheet, wbTmp As Workbook, fName As String
    sheets = Array("Calendar", "Products", "Customers", "Territories", "Sales_Raw", "Budget", "Forecast", "Market")

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the export folder"
        If .Show <> -1 Then Exit Sub
        folder = .SelectedItems(1)
    End With
    If Right$(folder, 1) <> "\" Then folder = folder & "\"

    On Error GoTo Fail
    SpeedOn
    For i = LBound(sheets) To UBound(sheets)
        Set ws = ThisWorkbook.Worksheets(sheets(i))
        Set wbTmp = Workbooks.Add(xlWBATWorksheet)
        ws.UsedRange.Copy
        wbTmp.Worksheets(1).Range("A1").PasteSpecial xlPasteValues
        wbTmp.Worksheets(1).Range("A1").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        fName = folder & "export_" & LCase$(sheets(i)) & ".csv"
        wbTmp.SaveAs Filename:=fName, FileFormat:=xlCSVUTF8, Local:=False
        wbTmp.Close SaveChanges:=False
    Next i
    WriteLog "ExportToCSV", (UBound(sheets) + 1) & " sheets exported to " & folder
    SpeedOff
    MsgBox (UBound(sheets) + 1) & " CSV files written to " & folder, vbInformation
    Exit Sub
Fail:
    If Not wbTmp Is Nothing Then wbTmp.Close SaveChanges:=False
    SpeedOff
    WriteLog "ExportToCSV", "FAILED: " & Err.Description
    MsgBox "Export failed: " & Err.Description, vbCritical
End Sub
