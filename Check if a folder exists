'function to check if a folder exists
Public Function FileFolderExists(byVal strFullPath As String) As Boolean
On Error GoTo ErrorHandler:
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then
        FileFolderExists = True
    End If
Exit Function
ErrorHandler:
    MsgBox err.description
    End
End Function
