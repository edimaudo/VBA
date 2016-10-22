Option Explicit

Public Sub aggregateWorksheets()
On Error GoTo ErrorHandler
    
    Dim currentSheet As Worksheet
    Dim pasteSheet As Worksheet
    Dim rangeString As String
    Dim currentSheetRowCount As Long
    Dim pasteRowCount As Long
    
    pasteRowCount = 2
    
    Application.ScreenUpdating = False
    Set pasteSheet = ThisWorkbook.Worksheets("Append all tabs")
    For Each currentSheet In ThisWorkbook.Worksheets
            currentSheetRowCount = currentSheet.Cells(Rows.Count, "B").End(xlUp).Row
            If currentSheetRowCount > 2 Then
                currentSheet.Range("A2:K" & CStr(currentSheetRowCount)).Copy
                pasteSheet.Range("A" & CStr(pasteRowCount) & ":K" & CStr(pasteRowCount + currentSheetRowCount)).PasteSpecial
                Application.CutCopyMode = False
            End If
            pasteRowCount = pasteSheet.Cells(Rows.Count, "B").End(xlUp).Row + 1      
    Next currentSheet
    
    Set currentSheet = Nothing
    Set pasteSheet = Nothing
    Application.ScreenUpdating = True
    MsgBox "Aggregation Complete"
Exit Sub
ErrorHandler:
    MsgBox Err.Description
    End
End Sub



Sub ClearDataInWorkSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Append all tabs")
    Application.ScreenUpdating = False
    ws.Rows("100000:100000").Select
    ws.Range(Selection, Selection.End(xlUp)).Select
    ws.Rows("2:100000").Select
    ws.Range("A100000").Activate
    Selection.Delete Shift:=xlUp
    Set ws = Nothing
End Sub
