Public Sub ClearRowContent(ByVal worksheetName as String)
	On Error GoTo errorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(worksheetName)
    Application.ScreenUpdating = False
    ws.Rows("2:100000").Delete Shift:=xlUp
    Set ws = Nothing
    Application.ScreenUpdating = True
    Exit sub
    errorHandler:
    	msgbox Err.Description
    	Resume Next
End Sub
