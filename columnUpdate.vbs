'loop through first column and only show the letters in column F

Option Explicit

Sub updateColumnF()

On Error GoTo ErrorHandler

Dim ws As Worksheet
Dim i As Long
Dim outputRow As Long
Dim rowEnd As Long

Set ws = ThisWorkbook.Sheets(1) 'Can also use sheet name e.g. "Sheet1"
rowEnd = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
outputRow = 1

For i = 1 To rowEnd
    If ws.Range("A" & i) <> "-" Then
        ws.Range("F" & outputRow).Value = ws.Range("A" & i).Value
        outputRow = outputRow + 1
    End If
Next i

Set ws = Nothing

MsgBox "Column F updated"

Exit Sub
ErrorHandler:
MsgBox "An error occured in the subrountine"

End Sub
