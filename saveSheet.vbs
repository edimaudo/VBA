'Save the excel file with datetime stamp
Sub SaveExcelSheet(ByVal filePath As String, ByVal fileName As String)
    fileName = fileName & "_" & CStr(Format(Date, "yyyy-mm-dd"))
    fileName = fileName & "_" & CStr(Format(Time, "hh-mm-ss AM/PM"))
    ActiveWorkbook.SaveAs filePath & fileName
    Application.Quit
End Sub
