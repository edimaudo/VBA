Option Explicit

Public Sub calculateAverage()
On Error GoTo ErrorHandler

Dim ws As Worksheet
Dim highest_avg As Double, lowest_avg As Double
Dim rowlength As Long
Dim dataArray(0 To 7) As Double
Dim dataArrayCounter As Long
Dim counter As Long
Dim startValue As Long
Dim sortedArray As Variant
Dim i As Long

Set ws = ThisWorkbook.Sheets(1) 'using the first worksheet
rowlength = Cells(Rows.Count, "A").End(xlUp).Row
dataArrayCounter = 0

startValue = rowlength - 7

For counter = startValue To rowlength
    dataArray(dataArrayCounter) = ws.Range("A" & counter).Value
    dataArrayCounter = dataArrayCounter + 1
Next counter

sortedArray = dataArray
sortedArray = BubbleSrt(sortedArray, True)


For i = LBound(sortedArray) To LBound(sortedArray) + 2
    lowest_avg = lowest_avg + sortedArray(i)
Next i

For i = UBound(sortedArray) - 2 To UBound(sortedArray)
    highest_avg = highest_avg + sortedArray(i) '
Next i

ws.Range("B" & 2).Value = lowest_avg
ws.Range("C" & 2).Value = highest_avg

Set ws = Nothing

Exit Sub
ErrorHandler:
    MsgBox "An error occured while calculating the average"
    End
End Sub

Public Function BubbleSrt(ArrayIn, Ascending As Boolean)

Dim SrtTemp As Variant
Dim i As Long
Dim j As Long


If Ascending = True Then
    For i = LBound(ArrayIn) To UBound(ArrayIn)
         For j = i + 1 To UBound(ArrayIn)
             If ArrayIn(i) > ArrayIn(j) Then
                 SrtTemp = ArrayIn(j)
                 ArrayIn(j) = ArrayIn(i)
                 ArrayIn(i) = SrtTemp
             End If
         Next j
     Next i
Else
    For i = LBound(ArrayIn) To UBound(ArrayIn)
         For j = i + 1 To UBound(ArrayIn)
             If ArrayIn(i) < ArrayIn(j) Then
                 SrtTemp = ArrayIn(j)
                 ArrayIn(j) = ArrayIn(i)
                 ArrayIn(i) = SrtTemp
             End If
         Next j
     Next i
End If

BubbleSrt = ArrayIn

End Function
