Public Sub SplitColumnA()
On Error GoTo errorHandler

Dim ws As Worksheet
Dim rowCount As Long
Dim count As Long
Dim rowRange As Range, cellRange As Range, splitRange As Range
Dim dataInfo As String
Dim outputData As Variant
Dim counter As Long
Dim output As String

Set ws = ThisWorkbook.Sheets("sheet1") 'update if needed
counter = 2
rowCount = ws.Cells(Rows.count, "A").End(xlUp).Row
Set rowRange = ws.Range("A2:A" & rowCount)

Application.ScreenUpdating = False

'replace ;
ws.Columns("A").Replace _
 What:=";", Replacement:=",", _
 SearchOrder:=xlByColumns, MatchCase:=True

'clean data
For count = rowCount To counter Step -1
    If Application.WorksheetFunction.IsNA(ws.Range("A" & count).Value) Then
        ws.Range("A" & count).EntireRow.Delete
    End If
Next count

count = 0

'split data
For Each cellRange In rowRange.Cells
        outputData = Split(cellRange.Value, ", ") 'update
        For count = 0 To UBound(outputData)
            ws.Range("B" & counter).Value = CStr(outputData(count))
            counter = counter + 1
        Next
Next


Set ws = Nothing

Application.ScreenUpdating = True

Exit Sub
errorHandler:
    MsgBox "An error occured while processing the data"
    
End Sub



