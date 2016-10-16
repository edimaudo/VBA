Option Compare Database
Option Explicit

'Works for MS Access
Sub StatusBar(ByVal tablename As String)
    On Error GoTo ErrorHandler
    Dim db As Database
    Dim rs As Recordset
    Dim rscounter As Long
    Dim retval As Variant
    Set db = CurrentDb
    Set rs = db.OpenRecordset(tablename)
    rs.MoveFirst
    rs.MoveLast
    rscounter = rs.RecordCount

    'Altering the Statusbar Property
    retval = SysCmd(4, Str(rscounter) & " tasks to run")

    If rscounter = 1 Then
        retval = SysCmd(5)
    End If

Set db = Nothing
Set rs = Nothing
Exit sub
ErrorHandler:
	msgbox "An error occured"
		end
End Sub
