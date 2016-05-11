Sub DeleteFile(ByVal pathinfo As String)
On Error GoTo ErrorHandler
    Kill pathinfo
Exit Sub
ErrorHandler:
If Err.Number = 53 Then
    Resume Next
End If
End Sub
