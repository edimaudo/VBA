Sub MergeFilesWithoutSpaces()
    Dim path As String, ThisWB As String, lngFilecounter As Long
    Dim wbDest As Workbook, shtDest As Worksheet, ws As Worksheet
    Dim Filename As String, Wkb As Workbook
    Dim CopyRng As Range, Dest As Range
    Dim RowofCopySheet As Integer
ThisWB = ActiveWorkbook.Name

path = "C:\UKSW CS Bom Expections\CS_BOM_Corrections\Archive"

RowofCopySheet = 2

Application.EnableEvents = False
Application.ScreenUpdating = False

Set shtDest = ActiveWorkbook.Sheets(1)
Filename = Dir(path & "\*.xls", vbNormal)
If Len(Filename) = 0 Then Exit Sub
Do Until Filename = vbNullString
    If Not Filename = ThisWB Then
        Set Wkb = Workbooks.Open(Filename:=path & "\" & Filename)
        Set CopyRng = Wkb.Sheets(1).Range(Cells(RowofCopySheet, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, Cells(1, Columns.Count).End(xlToLeft).Column))
        Set Dest = shtDest.Range("A" & shtDest.Cells(Rows.Count, 1).End(xlUp).Row + 1)
        CopyRng.Copy
        Dest.PasteSpecial xlPasteFormats
        Dest.PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False 'Clear Clipboard
        Wkb.Close False
    End If

    Filename = Dir()
Loop
