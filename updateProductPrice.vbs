Option Explicit

Public Sub updateProductPrice()
On Error GoTo ErrorHandler

Dim worksheet1 As Worksheet
Dim worksheet2 As Worksheet
Dim rowCounter As Long
Dim rowLength As Long
Dim rowInput As Long
Dim startValue As Long

Set worksheet1 = ThisWorkbook.Sheets(1)
Set worksheet2 = ThisWorkbook.Sheets(2)
rowCounter = 2
rowLength = Cells(Rows.Count, "A").End(xlUp).Row
rowInput = 2

For startValue = rowInput To rowLength
    If worksheet1.Range("C" & startValue).Value > 0 Then
        worksheet2.Range("A" & rowCounter).Value = worksheet1.Range("A" & startValue).Value
        worksheet2.Range("B" & rowCounter).Value = worksheet1.Range("B" & startValue).Value
        rowCounter = rowCounter + 1
    End If
Next startValue


Set worksheet1 = Nothing
Set worksheet2 = Nothing

Exit Sub
ErrorHandler:
    MsgBox "An error occured while updating product and price"
    End
End Sub
