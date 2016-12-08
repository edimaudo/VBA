Option Explicit

Sub Sample()
    Dim strFolder As String, strFile As String
    Dim MyData As String, strData() As String
    Dim FinalArray() As String
    Dim StartTime As String, endTime As String
    Dim n As Long, j As Long, i As Long

    strFolder = "C:\Temp\"

    strFile = Dir(strFolder & "*.csv")

    n = 0

    StartTime = Now

    Do While strFile <> ""
        Open strFolder & strFile For Binary As #1
        MyData = Space$(LOF(1))
        Get #1, , MyData
        Close #1

        strData() = Split(MyData, vbCrLf)
        ReDim Preserve FinalArray(j + UBound(strData) + 1)
        j = UBound(FinalArray)

        For i = LBound(strData) To UBound(strData)
            FinalArray(n) = strData(i)
            n = n + 1
        Next i

        strFile = Dir
    Loop

    endTime = Now

    Debug.Print "Process started at : " & StartTime
    Debug.Print "Process ended at : " & endTime
    Debug.Print UBound(FinalArray)
End Sub
