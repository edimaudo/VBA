Sub getAuthorData(ByVal folderpath As String)
On Error GoTo ErrorHandler
    Dim filename As String
    Dim fname As String
    Dim wkb As Workbook
    Dim firstspace As Long
    filename = Dir(folderpath & "*.xl*")
    Do Until filename = vbNullString
    Set wkb = Workbooks.Open(folderpath & filename, False, True)
    fname = wkb.BuiltinDocumentProperties("Last Author")
      msgbox fname
    wkb.Close
    filename = Dir()
    Loop
Exit Sub
ErrorHandler:
wkb.Close
    Resume Next
End Sub
