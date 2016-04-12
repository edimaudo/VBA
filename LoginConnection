Sub LoginConnection()
On Error GoTo ErrorHandler:
    Dim db As Database
    Dim ws As Workspace
   
    Dim ConnectionInfo As String


    ConnectionInfo = "" 'add connection string
	Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase("", False, True, ConnectionInfo)

    Set db = Nothing
    Set ws = Nothing
Exit Sub
ErrorHandler:
    MsgBox Err.Description
    End
End Sub
