Sub SCRIPTING_VLOOKUP(REF_TAB_WKSHT As String, REF_TAB_RNG As String, REF_TAB_KEY_COL As Integer, REF_TAB_OP_COL As Integer, INP_TAB_WKSHT As String, INP_TAB_INP_RNG As String, INP_TAB_INP_RNG_COL As Integer, OP_TAB_WKSHT As String, OP_TAB_OP_CELL As String)

    Dim dicLookupTable As Scripting.Dictionary
    Dim i As Long
    Dim sKey As String
    Dim vLookupValues As Variant
    Dim vLookupTable As Variant
    
    Set dicLookupTable = New Scripting.Dictionary
    dicLookupTable.CompareMode = vbTextCompare
    
    vLookupTable = Sheets(REF_TAB_WKSHT).Range(REF_TAB_RNG).Value
    For i = LBound(vLookupTable) To UBound(vLookupTable)
       sKey = vLookupTable(i, REF_TAB_KEY_COL)
       If Not dicLookupTable.Exists(sKey) Then _
          dicLookupTable(sKey) = vLookupTable(i, REF_TAB_OP_COL)
    Next i
       
    vLookupValues = Sheets(INP_TAB_WKSHT).Range(INP_TAB_INP_RNG)
       
    For i = LBound(vLookupValues) To UBound(vLookupValues)
       sKey = vLookupValues(i, INP_TAB_INP_RNG_COL)
                  
       If dicLookupTable.Exists(sKey) Then
          vLookupValues(i, INP_TAB_INP_RNG_COL) = dicLookupTable(sKey)
       Else
          vLookupValues(i, INP_TAB_INP_RNG_COL) = CVErr(xlErrNA)
       End If
    Next i
       
    Sheets(OP_TAB_WKSHT).Range(OP_TAB_OP_CELL).Resize(UBound(vLookupValues) - LBound(vLookupValues) + 1, 1) = vLookupValues
    
End Sub
