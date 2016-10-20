Sub SimpleCopyPaste()

On Error GoTo ErrorHandler:

Dim worksheet1 As Worksheet, worksheet2 As Worksheet

Set worksheet1 = ThisWorkbook.Worksheets("sheet1")
Set worksheet2 = ThisWorkbook.Worksheets("sheet2")

worksheet1.Columns("A:B").Copy
worksheet2.Columns("A:B").PasteSpecial
Application.CutCopyMode = False

Set worksheet1 = Nothing
Set worksheet2 = Nothing

Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub
