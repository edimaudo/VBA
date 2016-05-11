'Run batch file in VBA
'Make sure Microsoft Shell Controls and Automation is checked in the References in the VBA section (Tools -> References)
Sub RunBatch(ByVal filelocation As String, ByVal filename As String)
On Error GoTo ErrorHandler
    Call Shell(filelocation & filename, 1)
Exit Sub
ErrorHandler:
    MsgBox Err.Description
    Call log(Err.Description)
    End
End Sub
