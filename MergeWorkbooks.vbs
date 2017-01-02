Option Explicit
Public MyFiles As String

'To add the code to merge all workbooks in a folder
Sub MacMergeCode()
    Dim BaseWks As Worksheet
    Dim rnum As Long
    Dim CalcMode As Long
    Dim MySplit As Variant
    Dim FileInMyFiles As Long
    Dim Mybook As Workbook
    Dim sourceRange As Range
    Dim destrange As Range
    Dim SourceRcount As Long

    ActiveWindow.WindowState = xlNormal
    
    'Add a new workbook with one sheet
    Set BaseWks = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    BaseWks.Range("A1").Font.Size = 36
    BaseWks.Range("A1").Value = "Please Wait"
    rnum = 3

    'Change ScreenUpdating, Calculation and EnableEvents
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    'Clear MyFiles to be sure that it not return old info if no files are found
    MyFiles = ""

    'Get the files, set the level of folders and extension in the code line below
    Call GetFilesOnMacWithOrWithoutSubfolders(Level:=1, ExtChoice:=0, FileFilterOption:=0, FileNameFilterStr:="")
    'Level              :  1= Only the files in the folder you select, 2 to ? levels of subfolders
    'ExtChoice          :  0=(xls|xlsx|xlsm|xlsb), 1=xls , 2=xlsx, 3=xlsm, 4=xlsb, 5=csv, 6=txt, 7=all files, 8=(xlsx|xlsm|xlsb), 9=(csv|txt)
    'FileFilterOption   :  0=No Filter, 1=Begins, 2=Ends, 3=Contains
    'FileNameFilterStr  : Search string used when FileFilterOption = 1, 2 or 3

    ' Work with the files if MyFiles is not empty.
    If MyFiles <> "" Then

        MySplit = Split(MyFiles, Chr(10))
        For FileInMyFiles = LBound(MySplit) To UBound(MySplit) - 1

            Set Mybook = Nothing
            On Error Resume Next
            Set Mybook = Workbooks.Open(MySplit(FileInMyFiles))
            On Error GoTo 0

            If Not Mybook Is Nothing Then

                On Error Resume Next

                With Mybook.Worksheets(1)
                    Set sourceRange = .Range("A1:G10")
                End With

                If Err.Number > 0 Then
                    Err.Clear
                    Set sourceRange = Nothing
                Else
                    'if SourceRange use all columns then skip this file
                    If sourceRange.Columns.Count >= BaseWks.Columns.Count Then
                        Set sourceRange = Nothing
                    End If
                End If
                On Error GoTo 0

                If Not sourceRange Is Nothing Then

                    SourceRcount = sourceRange.Rows.Count

                    If rnum + SourceRcount >= BaseWks.Rows.Count Then
                        MsgBox "Sorry there are not enough rows in the sheet"
                        BaseWks.Columns.AutoFit
                        Mybook.Close savechanges:=False
                        GoTo ExitTheSub
                    Else

                        'Copy the file name in column A
                        With sourceRange
                            BaseWks.Cells(rnum, "A"). _
                                    Resize(.Rows.Count).Value = MySplit(FileInMyFiles)
                        End With

                        'Set the destrange
                        Set destrange = BaseWks.Range("B" & rnum)

                        'we copy the values from the sourceRange to the destrange
                        With sourceRange
                            Set destrange = destrange. _
                                            Resize(.Rows.Count, .Columns.Count)
                        End With
                        destrange.Value = sourceRange.Value

                        rnum = rnum + SourceRcount
                    End If
                End If
                Mybook.Close savechanges:=False
            End If

        Next FileInMyFiles
        BaseWks.Columns.AutoFit
    End If

ExitTheSub:
    BaseWks.Range("A1").Value = "Ready"
    'Restore ScreenUpdating, Calculation and EnableEvents
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub

'To merge data from columns in the source workbook to the target workbook
Sub MacMergeCodeColumns()
    Dim BaseWks As Worksheet
    Dim Cnum As Long
    Dim CalcMode As Long
    Dim MySplit As Variant
    Dim FileInMyFiles As Long
    Dim Mybook As Workbook
    Dim sourceRange As Range
    Dim destrange As Range
    Dim SourceCcount As Long

    ActiveWindow.WindowState = xlNormal

    'Add a new workbook with one sheet
    Set BaseWks = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    BaseWks.Range("A1").Font.Size = 36
    BaseWks.Range("A1").Value = "Please Wait"
    Cnum = 2

    'Change the ScreenUpdating, Calculation and EnableEvents settings
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    'Clear MyFiles so that it does not return the old data if no files are found.
    MyFiles = ""

    'Get the files, set the level of folders and extension in the code line below
    Call GetFilesOnMacWithOrWithoutSubfolders(Level:=1, ExtChoice:=0, FileFilterOption:=0, FileNameFilterStr:="")
    'Level              :  1= Only the files in the folder you select, 2 to ? levels of subfolders
    'ExtChoice          :  0=(xls|xlsx|xlsm|xlsb), 1=xls , 2=xlsx, 3=xlsm, 4=xlsb, 5=csv, 6=txt, 7=all files, 8=(xlsx|xlsm|xlsb), 9=(csv|txt)
    'FileFilterOption   :  0=No Filter, 1=Begins, 2=Ends, 3=Contains
    'FileNameFilterStr  : Search string used when FileFilterOption = 1, 2 or 3

    ' Work with the files if MyFiles is not empty.
    If MyFiles <> "" Then

        MySplit = Split(MyFiles, Chr(10))
        For FileInMyFiles = LBound(MySplit) To UBound(MySplit) - 1

            Set Mybook = Nothing
            On Error Resume Next
            Set Mybook = Workbooks.Open(MySplit(FileInMyFiles))
            On Error GoTo 0

            If Not Mybook Is Nothing Then

                On Error Resume Next
                
                With Mybook.Worksheets(1)
                    Set sourceRange = .Range("A1:A10")
                End With

                If Err.Number > 0 Then
                    Err.Clear
                    Set sourceRange = Nothing
                Else
                    'If the source range uses all of the rows
                    'then skip this file.
                    If sourceRange.Rows.Count >= BaseWks.Rows.Count Then
                        Set sourceRange = Nothing
                    End If
                End If
                On Error GoTo 0

                If Not sourceRange Is Nothing Then

                    SourceCcount = sourceRange.Columns.Count

                    If Cnum + SourceCcount >= BaseWks.Columns.Count Then
                        MsgBox "There are not enough columns in the sheet."
                        BaseWks.Columns.AutoFit
                        Mybook.Close savechanges:=False
                        GoTo ExitTheSub
                    Else

                        'Copy the file name in the first row.
                        With sourceRange
                            BaseWks.Cells(1, Cnum). _
                                    Resize(, .Columns.Count).Value = MySplit(FileInMyFiles)
                        End With

                        'Set the destination range.
                        Set destrange = BaseWks.Cells(2, Cnum)

                        'Copy the values from the source range
                        'to the destination range.
                        With sourceRange
                            Set destrange = destrange. _
                                            Resize(.Rows.Count, .Columns.Count)
                        End With
                        destrange.Value = sourceRange.Value

                        Cnum = Cnum + SourceCcount
                    End If
                End If
                Mybook.Close savechanges:=False
            End If

        Next FileInMyFiles
        BaseWks.Columns.AutoFit
    End If

ExitTheSub:
    BaseWks.Range("A1").Value = "Ready"
    'Restore the ScreenUpdating, Calculation and EnableEvents settings.
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub

'To add the custom functions to a standard module
Sub WINorMAC()
'Test for the current operating system.
    If Not Application.OperatingSystem Like "*Mac*" Then
        'OS is Windows
        Call My_Windows_Macro
    Else
        'OS is Mac so test to see if you are using Excel 
        '2011 or later.
        If Val(Application.Version) > 14 Then
            Call My_Mac_Macro
        End If
    End If
End Sub

Function RDB_Last(choice As Integer, rng As Range)
'Ron de Bruin, 5 May 2008
'Case 1 = last row
'Case 2 = last column
'Case 3 = last cell
    Dim lrw As Long
    Dim lcol As Integer

    Select Case choice

    Case 1:
        On Error Resume Next
        RDB_Last = rng.Find(What:="*", _
                            after:=rng.Cells(1), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
        On Error GoTo 0

    Case 2:
        On Error Resume Next
        RDB_Last = rng.Find(What:="*", _
                            after:=rng.Cells(1), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
        On Error GoTo 0

    Case 3:
        On Error Resume Next
        lrw = rng.Find(What:="*", _
                       after:=rng.Cells(1), _
                       Lookat:=xlPart, _
                       LookIn:=xlFormulas, _
                       SearchOrder:=xlByRows, _
                       SearchDirection:=xlPrevious, _
                       MatchCase:=False).Row
        On Error GoTo 0

        On Error Resume Next
        lcol = rng.Find(What:="*", _
                        after:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

        On Error Resume Next
        RDB_Last = rng.Parent.Cells(lrw, lcol).Address(False, False)
        If Err.Number > 0 Then
            RDB_Last = rng.Cells(1).Address(False, False)
            Err.Clear
        End If
        On Error GoTo 0
    End Select
End Function

Function GetFilesOnMacWithOrWithoutSubfolders(Level As Long, ExtChoice As Long, _
                                              FileFilterOption As Long, FileNameFilterStr As String)
'Ron de Bruin,Version 2 : 9 Nov 2012
'http://www.rondebruin.nl/mac.htm
'Thanks to DJ Bazzie Wazzie(poster on MacScripter) for his great help.
    Dim ScriptToRun As String
    Dim folderPath As String
    Dim FileNameFilter As String
    Dim Extensions As String

    On Error Resume Next
    folderPath = MacScript("choose folder as string")
    If folderPath = "" Then Exit Function
    On Error GoTo 0

    Select Case ExtChoice
    Case 0: Extensions = "(xls|xlsx|xlsm|xlsb)"  'xls, xlsx , xlsm, xlsb
    Case 1: Extensions = "xls"   'Only  xls
    Case 2: Extensions = "xlsx"  'Only xlsx
    Case 3: Extensions = "xlsm"  'Only xlsm
    Case 4: Extensions = "xlsb"  'Only xlsb
    Case 5: Extensions = "csv"   'Only csv
    Case 6: Extensions = "txt"   'Only txt
    Case 7: Extensions = ".*"    'All files with extension, use *.* for everything
    Case 8: Extensions = "(xlsx|xlsm|xlsb)"  'xlsx, xlsm , xlsb
    Case 9: Extensions = "(csv|txt)"   'csv and txt files
        'You can add more filter options if you want.
    End Select

    Select Case FileFilterOption
    Case 0: FileNameFilter = "'.*/[^~][^/]*\\." & Extensions & "$' "  'No Filter
    Case 1: FileNameFilter = "'.*/" & FileNameFilterStr & "[^~][^/]*\\." & Extensions & "$' "    'Begins with
    Case 2: FileNameFilter = "'.*/[^~][^/]*" & FileNameFilterStr & "\\." & Extensions & "$' "    ' Ends With
    Case 3: FileNameFilter = "'.*/([^~][^/]*" & FileNameFilterStr & "[^/]*|" & FileNameFilterStr & "[^/]*)\\." & Extensions & "$' "   'Contains
    End Select

    folderPath = MacScript("tell text 1 thru -2 of " & Chr(34) & folderPath & _
                           Chr(34) & " to return quoted form of it's POSIX Path")
    ScriptToRun = ScriptToRun & _
                  "set streamEditorCommand to " & _
                  Chr(34) & " |  tr  [/:] [:/] " & Chr(34) & Chr(13)
    ScriptToRun = ScriptToRun & _
                  "set streamEditorCommand to streamEditorCommand & " & _
                  Chr(34) & " | sed -e " & Chr(34) & "  & quoted form of (" & _
                  Chr(34) & " s.:." & Chr(34) & _
                "  & (POSIX file " & Chr(34) & "/" & Chr(34) & "  as string) & " & _
                  Chr(34) & "." & Chr(34) & " )" & Chr(13)
    ScriptToRun = ScriptToRun & "do shell script """ & "find -E " & _
                  folderPath & " -iregex " & FileNameFilter & "-maxdepth " & _
                  Level & """ & streamEditorCommand without altering line endings"

    On Error Resume Next
    MyFiles = MacScript(ScriptToRun)
    On Error GoTo 0
End Function
