Sub MoveExcelFile(byVal FromDir As String, byVal ToDir As String)
On Error GoTo ErrorHandler:

    Dim FSO As Object
    Dim strFile As String
    Set FSO = CreateObject("Scripting.FileSystemObject")

    strFile = Dir(FromDir & "*.xl*")
    While strFile <> ""
        If IsWorkBookOpen(FromDir & strFile) = True Then
            strFile = Dir()
        Else
            FSO.MoveFile Source:=FromDir & strFile, Destination:=ToDir
            strFile = Dir()
        End If
    Wend
    
    Set FSO = Nothing
Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
    Resume Next
End Sub
