Sub Select_File_Or_Files_Mac()
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFiles As String
    Dim MySplit As Variant
    Dim N As Long
    Dim Fname As String
    Dim mybook As Workbook

    On Error Resume Next
    MyPath = MacScript("return (path to documents folder) as String")
    'Or use MyPath = "Macintosh HD:Users:YourUserName:Desktop:TestFolder:"

    MyScript = "set applescript's text item delimiters to (ASCII character 10) " & vbNewLine & _
            "set theFiles to (choose file of type " & _
          " (""public.comma-separated-values-text"") " & _
            "with prompt ""Please select a file or files"" default location alias """ & _
            MyPath & """ multiple selections allowed true) as string" & vbNewLine & _
            "set applescript's text item delimiters to """" " & vbNewLine & _
            "return theFiles"

    MyFiles = MacScript(MyScript)
    On Error GoTo 0

    If MyFiles <> "" Then
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
        End With
    MySplit = Split(MyFiles, Chr(10))
        For N = LBound(MySplit) To UBound(MySplit)

            'Get file name only and test if it is open
            Fname = Right(MySplit(N), Len(MySplit(N)) - InStrRev(MySplit(N), _
            Application.PathSeparator, , 1))

                On Error Resume Next
                Set mybook = Workbooks.Open(MySplit(N))
                On Error GoTo 0
             Next

Worksheets("Rapport").Activate

With ActiveSheet.QueryTables.Add( _
        Connection:="TEXT;" & Fname, _
        Destination:=Range("A1"))
        .Name = "CSV" & Worksheets.Count + 1
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = xlMacintosh
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
End With

              End If

End Sub
