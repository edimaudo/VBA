Public Function IsAlphaBetical(TestString As String) As Boolean
     
    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
     
     'returns true if all characters in a string are alphabetical
     'returns false otherwise or for empty string
     
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[A-Za-z]" Then Exit Function
        Next
         
        IsAlphaBetical = True
    End If
     
End Function
 
Public Function IsAlphaNumeric(TestString As String) As Boolean
     
    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
     
     'returns true if all characters in a string are alphabetical
     '   or numeric
     'returns false otherwise or for empty string
     
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[0-9A-Za-z]" Then Exit Function
        Next
         
        IsAlphaNumeric = True
    End If
     
End Function
Public Function IsNumericOnly(TestString As String) As Boolean
     
    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
     
     'returns true if all characters in string are numeric
     'returns false otherwise or for empty string
     
     'this is different than VB's isNumeric
     'isNumeric returns true for something like 90.09
     'This function will return false
     
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[0-9]" Then Exit Function
        Next
         
        IsNumericOnly = True
    End If
     
End Function
