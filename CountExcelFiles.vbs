'Counts number of Excel files in a folder
Function CountFilesInFolder(byval folderpath As String) As Integer
    Dim path As String
    Dim count As Integer
    Dim filename As String
    path = folderpath & "*.xl*"
    filename = Dir(path)
    count = 0
    Do While filename <> "" 'While filename <> vbnullString
        count = count + 1
        filename = Dir()
    Loop
    CountFilesInFolder = count
End Function
