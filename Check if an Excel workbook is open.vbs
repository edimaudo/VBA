'Checks if an excel workbook is open
'filename should include filepath as well
Function IsWorkBookOpen(ByVal filename As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open filename For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
        Case 0:    IsWorkBookOpen = False
        Case 70:   IsWorkBookOpen = True
        Case Else: Error ErrNo
    End Select
End Function
