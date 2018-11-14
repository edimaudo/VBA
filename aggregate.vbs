
 

Option Explicit

 

Public Sub generate_level_analysis()

 

End Sub

 

 

'convert collection to array

Public Function toArray(col As Collection)

  Dim arr() As Variant

 

  If col.Count > 0 Then

 

    ReDim arr(1 To col.Count) As Variant

    Dim i As Variant

    For i = 1 To col.Count

      arr(i) = col(i)

    Next

   

  Else

    ReDim arr(0) As Variant

  End If

 

  toArray = arr

End Function

 

'calculate average

Public Function calculateAverage(ByVal arrayData As Variant) As Double

 

    Dim sumval As Double

    Dim countval As Double

    Dim i As Variant

    Dim calVal As Double

    sumval = 0

    countval = 0

   

    If IsEmpty(arrayData(1)) Then

        calVal = 0

    Else

        calVal = Application.WorksheetFunction.average(arrayData)

    End If

    '    Application.WorksheetFunction.average (arrayData)

   

    calculateAverage = calVal

 

End Function

 

'calculate count

Public Function calculateCount(ByVal arrayData As Variant) As Long

 

    Dim countval As Long

    Dim i As Long

   

    If IsEmpty(arrayData(1)) Then

        countval = 0

    Else

        countval = UBound(arrayData)

    End If

   

    calculateCount = countval

 

End Function

 

'calculate standard deviation

Public Function calculateStandardDeviation(ByVal arrayData As Variant) As Double

  Dim sumsqr As Double

  Dim i As Long

  Dim output As Double

 

  If IsEmpty(arrayData(1)) Then

    output = 0

  ElseIf UBound(arrayData) < 2 Then

    output = 0

  Else

    output = Application.WorksheetFunction.StDev(arrayData)

  End If

 

calculateStandardDeviation = output

End Function

 

 

'calculate confidence interval

Public Function calculateConfidenceInterval(ByVal arrayData As Variant, ByVal mean As Double, ByVal standard_deviation As Double) As Variant

 

    Dim significance As Double

 

    Dim confidenceValue As Double

    Dim upperConfidence As Double, lowerConfidence As Double

    Dim size As Double

    Dim output(0 To 1) As Double

   

    significance = 0.05

   

    If IsEmpty(arrayData(1)) Then

        output(0) = 0

        output(1) = 0

    ElseIf UBound(arrayData) < 2 Then

        output(0) = 0

        output(1) = 0

    Else

        size = UBound(arrayData)

        confidenceValue = Application.WorksheetFunction.Confidence(significance, standard_deviation, size)

        upperConfidence = mean + confidenceValue

        lowerConfidence = mean - confidenceValue

        output(0) = lowerConfidence

        output(1) = upperConfidence

    End If

 

calculateConfidenceInterval = output

 

End Function

 

Public Function removeOutliers(ByVal collectionData As Collection) As Variant

 

    Dim newcol As New Collection

    Dim lowerQuartile As Double, upperQuartile As Double, InterQuartile As Double

    Dim sumval As Double

    Dim countval As Double

    Dim calVal As Double

    Dim i As Long

    Dim lowerQuartileBound As Double, upperQuartileBound As Double

   

    Dim finalArray As Variant

    Dim tempArray As Variant

   

    

    If collectionData.Count > 0 Then

        tempArray = toArray(collectionData)

        lowerQuartile = Application.WorksheetFunction.Quartile(tempArray, 1)

        upperQuartile = Application.WorksheetFunction.Quartile(tempArray, 3)

        InterQuartile = upperQuartile - lowerQuartile

       

        lowerQuartileBound = lowerQuartile - (1.5 * InterQuartile)

        upperQuartileBound = upperQuartile + (1.5 * InterQuartile)

       

        For i = LBound(tempArray) To UBound(tempArray)

            If tempArray(i) >= lowerQuartileBound And tempArray(i) <= upperQuartileBound Then

                newcol.Add (tempArray(i))

            End If

        Next i

       

    End If

       

    finalArray = toArray(newcol)

   

    removeOutliers = finalArray

   

End Function

 

 

Public Sub calculate_level_time()

 

On Error GoTo errorhandler:

 

    Dim masterData As Worksheet, levels As Worksheet, calcSheet As Worksheet 'worksheets

   

    Dim masterDatarowcount As Long, calcSheetrowcount As Long, levelrowcount As Long 'row count

   

    Dim l1value As String, l2value As String, l3value As String, l4value As String 'level information

   

    Dim l As Long, m As Long 'counters

   

    Dim averageinfo(0 To 1) As Double

    Dim stddevinfo(0 To 1) As Double

    Dim confidenceIntervalinfo(0 To 3) As Double

    Dim countinfo(0 To 1) As Long

   

    Dim arrayDataInfo As Variant

    Dim arrayDataInfoNoOutliers As Variant

    Dim timeInfo As New Collection

   

    Dim leveldata(1 To 5) As Variant

    Dim calc(1 To 10) As Double

   

    Dim calcRange As String

    Dim leveldataRange As String

   

    

    Set masterData = ThisWorkbook.Sheets("Master Data")

    Set levels = ThisWorkbook.Sheets("Levels")

    Set calcSheet = ThisWorkbook.Sheets("Calc_Time_Levels")


    'clear old data

    Application.ScreenUpdating = False

    calcSheet.Rows("2:10000").Delete Shift:=xlUp

   

    masterDatarowcount = masterData.Cells(masterData.Rows.Count, "D").End(xlUp).Row

    calcSheetrowcount = 2

    levelrowcount = levels.Cells(levels.Rows.Count, "A").End(xlUp).Row

   

    For l = 2 To levelrowcount

        l1value = levels.Range("A" & l).Value

        l2value = levels.Range("B" & l).Value

        l3value = levels.Range("C" & l).Value

        l4value = levels.Range("D" & l).Value

       

        leveldata(1) = l1value

        leveldata(2) = l2value

        leveldata(3) = l3value

        leveldata(4) = l4value

       

        

        

        For m = 2 To masterDatarowcount

            If masterData.Range("U" & m).Value = l1value And _

               masterData.Range("V" & m).Value = l2value And _

               masterData.Range("W" & m).Value = l3value And _

               masterData.Range("X" & m).Value = l4value Then

              

                    timeInfo.Add masterData.Range("AL" & m).Value

       

            End If

        Next m

       

        leveldataRange = "A" & CStr(calcSheetrowcount) & ":" & "D" & CStr(calcSheetrowcount)

        calcRange = "E" & CStr(calcSheetrowcount) & ":" & "N" & CStr(calcSheetrowcount)

       

 

       

        'calculate statistics

        arrayDataInfoNoOutliers = removeOutliers(timeInfo)

        arrayDataInfo = toArray(timeInfo)

       

        averageinfo(0) = calculateAverage(arrayDataInfoNoOutliers)

        averageinfo(1) = calculateAverage(arrayDataInfo)

        stddevinfo(0) = calculateStandardDeviation(arrayDataInfoNoOutliers)

        stddevinfo(1) = calculateStandardDeviation(arrayDataInfo)

        confidenceIntervalinfo(0) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(0)

        confidenceIntervalinfo(1) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(1)

        confidenceIntervalinfo(2) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(0)

        confidenceIntervalinfo(3) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(1)

        countinfo(0) = calculateCount(arrayDataInfoNoOutliers)

        countinfo(1) = calculateCount(arrayDataInfo)

       

        

        'add calculations

        calc(1) = averageinfo(0)

        calc(2) = averageinfo(1)

        calc(3) = stddevinfo(0)

        calc(4) = stddevinfo(1)

        calc(5) = confidenceIntervalinfo(0)

        calc(6) = confidenceIntervalinfo(1)

        calc(7) = confidenceIntervalinfo(2)

        calc(8) = confidenceIntervalinfo(3)

        calc(9) = countinfo(0)

        calc(10) = countinfo(1)

       

        

        'output

        calcSheet.Range(leveldataRange).Value = leveldata

        calcSheet.Range(calcRange).Value = calc

       

        Set timeInfo = New Collection

       

        calcSheetrowcount = calcSheetrowcount + 1

       

        

                

    Next l

   

    'clear objects

    Set masterData = Nothing

    Set levels = Nothing

    Set calcSheet = Nothing

   

    

    MsgBox "levels Complete"

   

    Application.ScreenUpdating = True


 

Exit Sub

errorhandler:

    MsgBox Err.Description

    End

 

End Sub

 

 

'calculate store level calculation

 

Public Sub calculate_store_level()

 

On Error GoTo errorhandler

 

 

    Dim masterData As Worksheet, levels As Worksheet, calcSheet As Worksheet, store As Worksheet 'worksheets

   

    Dim masterDatarowcount As Long, calcSheetrowcount As Long, levelrowcount As Long, storerowcount As Long 'row count

   

    Dim l1value As String, l2value As String, l3value As String, l4value As String, l5value As String 'level information

   

    Dim storenumber As String

   

    Dim l As Long, m As Long, s As Long 'counters

   

    Dim averageinfo(0 To 1) As Double

    Dim stddevinfo(0 To 1) As Double

    Dim confidenceIntervalinfo(0 To 3) As Double

    Dim countinfo(0 To 1) As Long

   

    Dim arrayDataInfo As Variant

    Dim arrayDataInfoNoOutliers As Variant

    Dim timeInfo As New Collection

   

    Dim leveldata(1 To 5) As Variant

    Dim calc(1 To 10) As Double

   

    Dim calcRange As String

    Dim leveldataRange As String

   

    

    Set masterData = ThisWorkbook.Sheets("Master Data")

    Set levels = ThisWorkbook.Sheets("store_Level")

    Set calcSheet = ThisWorkbook.Sheets("Calc_Time_Store")

    Set store = ThisWorkbook.Sheets("Store")


    'clear old data

    Application.ScreenUpdating = False

    calcSheet.Rows("2:100000").Delete Shift:=xlUp

   

    masterDatarowcount = masterData.Cells(masterData.Rows.Count, "D").End(xlUp).Row

    calcSheetrowcount = 2

    levelrowcount = levels.Cells(levels.Rows.Count, "A").End(xlUp).Row

    storerowcount = store.Cells(store.Rows.Count, "A").End(xlUp).Row

   

    

            For l = 2 To levelrowcount

                l1value = levels.Range("A" & l).Value

                l2value = levels.Range("B" & l).Value

                l3value = levels.Range("C" & l).Value

                l4value = levels.Range("D" & l).Value

                l5value = levels.Range("E" & l).Value

               

                leveldata(1) = l1value

                leveldata(2) = l2value

                leveldata(3) = l3value

                leveldata(4) = l4value

                leveldata(5) = l5value

               

                

                For m = 2 To masterDatarowcount

                    If masterData.Range("F" & m).Value = l1value And _

                       masterData.Range("U" & m).Value = l2value And _

                       masterData.Range("V" & m).Value = l3value And _

                       masterData.Range("W" & m).Value = l4value And _

                       masterData.Range("X" & m).Value = l5value Then

                            timeInfo.Add masterData.Range("AL" & m).Value

               

                    End If

                Next m

               

                leveldataRange = "A" & CStr(calcSheetrowcount) & ":" & "E" & CStr(calcSheetrowcount)

                calcRange = "F" & CStr(calcSheetrowcount) & ":" & "O" & CStr(calcSheetrowcount)

               

        

                

                'calculate statistics

                arrayDataInfoNoOutliers = removeOutliers(timeInfo)

                arrayDataInfo = toArray(timeInfo)

               

                averageinfo(0) = calculateAverage(arrayDataInfoNoOutliers)

                averageinfo(1) = calculateAverage(arrayDataInfo)

                stddevinfo(0) = calculateStandardDeviation(arrayDataInfoNoOutliers)

                stddevinfo(1) = calculateStandardDeviation(arrayDataInfo)

                confidenceIntervalinfo(0) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(0)

                confidenceIntervalinfo(1) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(1)

                confidenceIntervalinfo(2) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(0)

                confidenceIntervalinfo(3) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(1)

                countinfo(0) = calculateCount(arrayDataInfoNoOutliers)

                countinfo(1) = calculateCount(arrayDataInfo)

               

                

                'add calculations

                calc(1) = averageinfo(0)

                calc(2) = averageinfo(1)

                calc(3) = stddevinfo(0)

                calc(4) = stddevinfo(1)

                calc(5) = confidenceIntervalinfo(0)

                calc(6) = confidenceIntervalinfo(1)

                calc(7) = confidenceIntervalinfo(2)

                calc(8) = confidenceIntervalinfo(3)

                calc(9) = countinfo(0)

                calc(10) = countinfo(1)

               

                

                'output

                calcSheet.Range(leveldataRange).Value = leveldata

                calcSheet.Range(calcRange).Value = calc

               

                Set timeInfo = New Collection

               

                calcSheetrowcount = calcSheetrowcount + 1

                  

            Next l

   

    

    'clear objects

    Set masterData = Nothing

    Set levels = Nothing

    Set calcSheet = Nothing

    Set store = Nothing

   

    MsgBox "store Complete"

   

    Application.ScreenUpdating = True

 

 

Exit Sub

errorhandler:

    MsgBox Err.Description

    End

 

End Sub

 

 

'add check for day

 

Public Sub calculate_day_level()

 

On Error GoTo errorhandler

 

    Dim masterData As Worksheet, levels As Worksheet, calcSheet As Worksheet, day As Worksheet 'worksheets

   

    Dim masterDatarowcount As Long, calcSheetrowcount As Long, levelrowcount As Long, dayrowcount As Long 'row count

   

    Dim l1value As String, l2value As String, l3value As String, l4value As String, l5value As String 'level information

   

    Dim l As Long, m As Long, s As Long 'counters

   

    Dim averageinfo(0 To 1) As Double

    Dim stddevinfo(0 To 1) As Double

    Dim confidenceIntervalinfo(0 To 3) As Double

    Dim countinfo(0 To 1) As Long

   

    Dim arrayDataInfo As Variant

    Dim arrayDataInfoNoOutliers As Variant

    Dim timeInfo As New Collection

   

    Dim leveldata(1 To 5) As Variant

    Dim calc(1 To 10) As Double

   

    Dim calcRange As String

    Dim leveldataRange As String

   

    

    Set masterData = ThisWorkbook.Sheets("Master Data")

    Set levels = ThisWorkbook.Sheets("day_Level")

    Set calcSheet = ThisWorkbook.Sheets("Calc_Time_Day")

    Set day = ThisWorkbook.Sheets("day")


    'clear old data

    Application.ScreenUpdating = False

    calcSheet.Rows("2:100000").Delete Shift:=xlUp

   

    masterDatarowcount = masterData.Cells(masterData.Rows.Count, "D").End(xlUp).Row

    calcSheetrowcount = 2

    levelrowcount = levels.Cells(levels.Rows.Count, "A").End(xlUp).Row

    dayrowcount = day.Cells(day.Rows.Count, "A").End(xlUp).Row

   

    

            For l = 2 To levelrowcount

                l1value = levels.Range("A" & l).Value

                l2value = levels.Range("B" & l).Value

                l3value = levels.Range("C" & l).Value

                l4value = levels.Range("D" & l).Value

                l5value = levels.Range("E" & l).Value

               

                leveldata(1) = l1value

                leveldata(2) = l2value

                leveldata(3) = l3value

                leveldata(4) = l4value

                leveldata(5) = l5value

               

                

                For m = 2 To masterDatarowcount

                    If masterData.Range("C" & m).Value = l1value And _

                       masterData.Range("U" & m).Value = l2value And _

                       masterData.Range("V" & m).Value = l3value And _

                       masterData.Range("W" & m).Value = l4value And _

                       masterData.Range("X" & m).Value = l5value Then

 

                            timeInfo.Add masterData.Range("AL" & m).Value

               

                    End If

                Next m

               

                leveldataRange = "A" & CStr(calcSheetrowcount) & ":" & "E" & CStr(calcSheetrowcount)

                calcRange = "F" & CStr(calcSheetrowcount) & ":" & "O" & CStr(calcSheetrowcount)

               

                'calculate statistics

                arrayDataInfoNoOutliers = removeOutliers(timeInfo)

                arrayDataInfo = toArray(timeInfo)

               

                averageinfo(0) = calculateAverage(arrayDataInfoNoOutliers)

                averageinfo(1) = calculateAverage(arrayDataInfo)

                stddevinfo(0) = calculateStandardDeviation(arrayDataInfoNoOutliers)

                stddevinfo(1) = calculateStandardDeviation(arrayDataInfo)

                confidenceIntervalinfo(0) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(0)

                confidenceIntervalinfo(1) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(1)

                confidenceIntervalinfo(2) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(0)

                confidenceIntervalinfo(3) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(1)

                countinfo(0) = calculateCount(arrayDataInfoNoOutliers)

                countinfo(1) = calculateCount(arrayDataInfo)

               

                

                'add calculations

                calc(1) = averageinfo(0)

                calc(2) = averageinfo(1)

                calc(3) = stddevinfo(0)

                calc(4) = stddevinfo(1)

                calc(5) = confidenceIntervalinfo(0)

                calc(6) = confidenceIntervalinfo(1)

                calc(7) = confidenceIntervalinfo(2)

                calc(8) = confidenceIntervalinfo(3)

                calc(9) = countinfo(0)

                calc(10) = countinfo(1)

               

                

                'output

                calcSheet.Range(leveldataRange).Value = leveldata

                calcSheet.Range(calcRange).Value = calc

               

                Set timeInfo = New Collection

               

                calcSheetrowcount = calcSheetrowcount + 1

                  

            Next l

   

    

    'clear objects

    Set masterData = Nothing

    Set levels = Nothing

    Set calcSheet = Nothing

    Set day = Nothing

   

    MsgBox "Day Complete"

   

    Application.ScreenUpdating = True

 

 

Exit Sub

errorhandler:

    MsgBox Err.Description

    End

 

End Sub

 

 

'add check for day-store

 

 

Public Sub calculate_level_level1()

 

On Error GoTo errorhandler

 

    Dim masterData As Worksheet, levels As Worksheet, calcSheet As Worksheet 'worksheets

   

    Dim masterDatarowcount As Long, calcSheetrowcount As Long, levelrowcount As Long, dayrowcount As Long 'row count

   

    Dim l1value As String, l2value As String, l3value As String, l4value As String, l5value As String 'level information

   

    Dim l As Long, m As Long, s As Long 'counters

   

    Dim averageinfo(0 To 1) As Double

    Dim stddevinfo(0 To 1) As Double

    Dim confidenceIntervalinfo(0 To 3) As Double

    Dim countinfo(0 To 1) As Long

   

    Dim arrayDataInfo As Variant

    Dim arrayDataInfoNoOutliers As Variant

    Dim timeInfo As New Collection

   

    Dim leveldata(1 To 5) As Variant

    Dim calc(1 To 10) As Double

   

    Dim calcRange As String

    Dim leveldataRange As String

   

    

    Set masterData = ThisWorkbook.Sheets("Master Data")

    Set levels = ThisWorkbook.Sheets("Level_1")

    Set calcSheet = ThisWorkbook.Sheets("Calc_Time_Level1")

   

 

    'clear old data

    Application.ScreenUpdating = False

    calcSheet.Rows("2:100000").Delete Shift:=xlUp

   

    masterDatarowcount = masterData.Cells(masterData.Rows.Count, "D").End(xlUp).Row

    calcSheetrowcount = 2

    levelrowcount = levels.Cells(levels.Rows.Count, "A").End(xlUp).Row

    'dayrowcount = day.Cells(day.Rows.Count, "A").End(xlUp).Row

   

    

            For l = 2 To levelrowcount

                l1value = levels.Range("A" & l).Value

                'l2value = levels.Range("B" & l).Value

                'l3value = levels.Range("C" & l).Value

                'l4value = levels.Range("D" & l).Value

                'l5value = levels.Range("E" & l).Value

               

                leveldata(1) = l1value

                'leveldata(2) = l2value

                'leveldata(3) = l3value

                'leveldata(4) = l4value

                'leveldata(5) = l5value

               

                

                For m = 2 To masterDatarowcount

                    If masterData.Range("U" & m).Value = l1value Then

                            timeInfo.Add masterData.Range("AL" & m).Value

                

                    End If

                Next m

               

                leveldataRange = "A" & CStr(calcSheetrowcount) & ":" & "A" & CStr(calcSheetrowcount)

                calcRange = "B" & CStr(calcSheetrowcount) & ":" & "K" & CStr(calcSheetrowcount)

               

                'calculate statistics

                arrayDataInfoNoOutliers = removeOutliers(timeInfo)

                arrayDataInfo = toArray(timeInfo)

               

                averageinfo(0) = calculateAverage(arrayDataInfoNoOutliers)

                averageinfo(1) = calculateAverage(arrayDataInfo)

                stddevinfo(0) = calculateStandardDeviation(arrayDataInfoNoOutliers)

                stddevinfo(1) = calculateStandardDeviation(arrayDataInfo)

                confidenceIntervalinfo(0) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(0)

                confidenceIntervalinfo(1) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(1)

                confidenceIntervalinfo(2) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(0)

                confidenceIntervalinfo(3) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(1)

                countinfo(0) = calculateCount(arrayDataInfoNoOutliers)

                countinfo(1) = calculateCount(arrayDataInfo)

               

                

                'add calculations

                calc(1) = averageinfo(0)

                calc(2) = averageinfo(1)

                calc(3) = stddevinfo(0)

                calc(4) = stddevinfo(1)

                calc(5) = confidenceIntervalinfo(0)

                calc(6) = confidenceIntervalinfo(1)

                calc(7) = confidenceIntervalinfo(2)

                calc(8) = confidenceIntervalinfo(3)

                calc(9) = countinfo(0)

                calc(10) = countinfo(1)

               

                

                'output

                calcSheet.Range(leveldataRange).Value = leveldata(1)

                calcSheet.Range(calcRange).Value = calc

               

                Set timeInfo = New Collection

               

                calcSheetrowcount = calcSheetrowcount + 1

                  

            Next l

   

    

    'clear objects

    Set masterData = Nothing

    Set levels = Nothing

    Set calcSheet = Nothing

   

    MsgBox "Level 1 Complete"

   

    Application.ScreenUpdating = True

 

 

Exit Sub

errorhandler:

    MsgBox Err.Description

    End

 

End Sub

 

 

Public Sub calculate_level_level1_2()

 

On Error GoTo errorhandler

 

    Dim masterData As Worksheet, levels As Worksheet, calcSheet As Worksheet 'worksheets

   

    Dim masterDatarowcount As Long, calcSheetrowcount As Long, levelrowcount As Long, dayrowcount As Long 'row count

   

    Dim l1value As String, l2value As String, l3value As String, l4value As String, l5value As String 'level information

   

    Dim l As Long, m As Long, s As Long 'counters

   

    Dim averageinfo(0 To 1) As Double

    Dim stddevinfo(0 To 1) As Double

    Dim confidenceIntervalinfo(0 To 3) As Double

    Dim countinfo(0 To 1) As Long

   

    Dim arrayDataInfo As Variant

    Dim arrayDataInfoNoOutliers As Variant

    Dim timeInfo As New Collection

   

    Dim leveldata(1 To 2) As Variant

    Dim calc(1 To 10) As Double

   

    Dim calcRange As String

    Dim leveldataRange As String

   

    

    Set masterData = ThisWorkbook.Sheets("Master Data")

    Set levels = ThisWorkbook.Sheets("Level_1_2")

    Set calcSheet = ThisWorkbook.Sheets("Calc_Time_Level1_2")

   

 

    'clear old data

    Application.ScreenUpdating = False

    calcSheet.Rows("2:100000").Delete Shift:=xlUp

   

    masterDatarowcount = masterData.Cells(masterData.Rows.Count, "D").End(xlUp).Row

    calcSheetrowcount = 2

    levelrowcount = levels.Cells(levels.Rows.Count, "A").End(xlUp).Row

 

   

            For l = 2 To levelrowcount

                l1value = levels.Range("A" & l).Value

                l2value = levels.Range("B" & l).Value

                'l3value = levels.Range("C" & l).Value

                'l4value = levels.Range("D" & l).Value

                'l5value = levels.Range("E" & l).Value

                

                leveldata(1) = l1value

                leveldata(2) = l2value

                'leveldata(3) = l3value

                'leveldata(4) = l4value

                'leveldata(5) = l5value

               

                

                For m = 2 To masterDatarowcount

                    If masterData.Range("U" & m).Value = l1value And masterData.Range("V" & m).Value = l2value Then

                            timeInfo.Add masterData.Range("AL" & m).Value

                    End If

                Next m

               

                leveldataRange = "A" & CStr(calcSheetrowcount) & ":" & "B" & CStr(calcSheetrowcount)

                calcRange = "C" & CStr(calcSheetrowcount) & ":" & "L" & CStr(calcSheetrowcount)

               

                'calculate statistics

                arrayDataInfoNoOutliers = removeOutliers(timeInfo)

                arrayDataInfo = toArray(timeInfo)

               

                averageinfo(0) = calculateAverage(arrayDataInfoNoOutliers)

                averageinfo(1) = calculateAverage(arrayDataInfo)

                stddevinfo(0) = calculateStandardDeviation(arrayDataInfoNoOutliers)

                stddevinfo(1) = calculateStandardDeviation(arrayDataInfo)

                confidenceIntervalinfo(0) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(0)

                confidenceIntervalinfo(1) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(1)

                confidenceIntervalinfo(2) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(0)

                confidenceIntervalinfo(3) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(1)

                countinfo(0) = calculateCount(arrayDataInfoNoOutliers)

                countinfo(1) = calculateCount(arrayDataInfo)

               

                

                'add calculations

                calc(1) = averageinfo(0)

                calc(2) = averageinfo(1)

                calc(3) = stddevinfo(0)

                calc(4) = stddevinfo(1)

                calc(5) = confidenceIntervalinfo(0)

                calc(6) = confidenceIntervalinfo(1)

                calc(7) = confidenceIntervalinfo(2)

                calc(8) = confidenceIntervalinfo(3)

                calc(9) = countinfo(0)

                calc(10) = countinfo(1)

               

                

                'output

                calcSheet.Range(leveldataRange).Value = leveldata

                calcSheet.Range(calcRange).Value = calc

               

                Set timeInfo = New Collection

               

                calcSheetrowcount = calcSheetrowcount + 1

                  

            Next l

   

    

    'clear objects

    Set masterData = Nothing

    Set levels = Nothing

    Set calcSheet = Nothing

   

    MsgBox "Level 1 & 2 Complete"

   

    Application.ScreenUpdating = True

 

 

Exit Sub

errorhandler:

    MsgBox Err.Description

    End

 

End Sub

 

 

 

Public Sub calculate_level_level1_2_3()

 

On Error GoTo errorhandler

 

    Dim masterData As Worksheet, levels As Worksheet, calcSheet As Worksheet 'worksheets

   

    Dim masterDatarowcount As Long, calcSheetrowcount As Long, levelrowcount As Long, dayrowcount As Long 'row count

   

    Dim l1value As String, l2value As String, l3value As String, l4value As String, l5value As String 'level information

   

    Dim l As Long, m As Long, s As Long 'counters

   

    Dim averageinfo(0 To 1) As Double

    Dim stddevinfo(0 To 1) As Double

    Dim confidenceIntervalinfo(0 To 3) As Double

    Dim countinfo(0 To 1) As Long

   

    Dim arrayDataInfo As Variant

    Dim arrayDataInfoNoOutliers As Variant

    Dim timeInfo As New Collection

   

    Dim leveldata(1 To 3) As Variant

    Dim calc(1 To 10) As Double

   

    Dim calcRange As String

    Dim leveldataRange As String

   

    

    Set masterData = ThisWorkbook.Sheets("Master Data")

    Set levels = ThisWorkbook.Sheets("Level_1_2_3")

    Set calcSheet = ThisWorkbook.Sheets("Calc_Time_Level1_2_3")

   

 

    'clear old data

    Application.ScreenUpdating = False

    calcSheet.Rows("2:100000").Delete Shift:=xlUp

   

    masterDatarowcount = masterData.Cells(masterData.Rows.Count, "D").End(xlUp).Row

    calcSheetrowcount = 2

    levelrowcount = levels.Cells(levels.Rows.Count, "A").End(xlUp).Row

 

   

            For l = 2 To levelrowcount

                l1value = levels.Range("A" & l).Value

                l2value = levels.Range("B" & l).Value

                l3value = levels.Range("C" & l).Value

                'l4value = levels.Range("D" & l).Value

                'l5value = levels.Range("E" & l).Value

               

                leveldata(1) = l1value

                leveldata(2) = l2value

                leveldata(3) = l3value

                'leveldata(4) = l4value

                'leveldata(5) = l5value

               

                

                For m = 2 To masterDatarowcount

                    If masterData.Range("U" & m).Value = l1value And masterData.Range("V" & m).Value = l2value And _

                        masterData.Range("W" & m).Value = l3value Then

                            timeInfo.Add masterData.Range("AL" & m).Value

                    End If

                Next m

               

                leveldataRange = "A" & CStr(calcSheetrowcount) & ":" & "C" & CStr(calcSheetrowcount)

                calcRange = "D" & CStr(calcSheetrowcount) & ":" & "M" & CStr(calcSheetrowcount)

               

                'calculate statistics

                arrayDataInfoNoOutliers = removeOutliers(timeInfo)

                arrayDataInfo = toArray(timeInfo)

               

                averageinfo(0) = calculateAverage(arrayDataInfoNoOutliers)

                averageinfo(1) = calculateAverage(arrayDataInfo)

                stddevinfo(0) = calculateStandardDeviation(arrayDataInfoNoOutliers)

                stddevinfo(1) = calculateStandardDeviation(arrayDataInfo)

                confidenceIntervalinfo(0) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(0)

                confidenceIntervalinfo(1) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(1)

                confidenceIntervalinfo(2) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(0)

                confidenceIntervalinfo(3) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(1)

                countinfo(0) = calculateCount(arrayDataInfoNoOutliers)

                countinfo(1) = calculateCount(arrayDataInfo)

               

                

                'add calculations

                calc(1) = averageinfo(0)

                calc(2) = averageinfo(1)

                calc(3) = stddevinfo(0)

                calc(4) = stddevinfo(1)

                calc(5) = confidenceIntervalinfo(0)

                calc(6) = confidenceIntervalinfo(1)

                calc(7) = confidenceIntervalinfo(2)

                calc(8) = confidenceIntervalinfo(3)

                calc(9) = countinfo(0)

                calc(10) = countinfo(1)

               

                

                'output

                calcSheet.Range(leveldataRange).Value = leveldata

                calcSheet.Range(calcRange).Value = calc

               

                Set timeInfo = New Collection

               

                calcSheetrowcount = calcSheetrowcount + 1

                  

            Next l

    

    

    'clear objects

    Set masterData = Nothing

    Set levels = Nothing

    Set calcSheet = Nothing

   

    MsgBox "Level 1 to 3 Complete"

   

    Application.ScreenUpdating = True

 

 

Exit Sub

errorhandler:

    MsgBox Err.Description

    End

 

End Sub

 

 

Public Sub driver_level()

 

On Error GoTo errorhandler

 

    Dim masterData As Worksheet, levels As Worksheet, calcSheet As Worksheet 'worksheets

   

    Dim masterDatarowcount As Long, calcSheetrowcount As Long, levelrowcount As Long, dayrowcount As Long 'row count

   

    Dim l1value As String, l2value As String, l3value As String, l4value As String, l5value As String, l6value As String 'level information

   

    Dim l As Long, m As Long, s As Long 'counters

   

    Dim averageinfo(0 To 1) As Double

    Dim stddevinfo(0 To 1) As Double

    Dim confidenceIntervalinfo(0 To 3) As Double

    Dim countinfo(0 To 1) As Long

   

    Dim arrayDataInfo As Variant

    Dim arrayDataInfoNoOutliers As Variant

    Dim timeInfo As New Collection

   

    Dim leveldata(1 To 6) As Variant

    Dim calc(1 To 12) As Double

   

    Dim calcRange As String

    Dim leveldataRange As String

    Dim totalcountpercent(0 To 1) As Double

   

    

    Set masterData = ThisWorkbook.Sheets("Master Data")

    Set levels = ThisWorkbook.Sheets("High_Low_levels")

    Set calcSheet = ThisWorkbook.Sheets("Calc_Time_High_Low")

   

 

    'clear old data

    Application.ScreenUpdating = False

    calcSheet.Rows("2:100000").Delete Shift:=xlUp

   

    masterDatarowcount = masterData.Cells(masterData.Rows.Count, "D").End(xlUp).Row

    calcSheetrowcount = 2

    levelrowcount = levels.Cells(levels.Rows.Count, "A").End(xlUp).Row

 

   

            For l = 2 To levelrowcount

                l1value = levels.Range("A" & l).Value

                l2value = levels.Range("B" & l).Value

                l3value = levels.Range("C" & l).Value

                l4value = levels.Range("D" & l).Value

                l5value = levels.Range("E" & l).Value

                l6value = levels.Range("F" & l).Value

               

                leveldata(1) = l1value

                leveldata(2) = l2value

                leveldata(3) = l3value

                leveldata(4) = l4value

                leveldata(5) = l5value

                leveldata(6) = l6value

               

                For m = 2 To masterDatarowcount

                    If masterData.Range("U" & m).Value = l1value And masterData.Range("V" & m).Value = l2value And _

                        masterData.Range("W" & m).Value = l3value And masterData.Range("X" & m).Value = l4value And _

                        masterData.Range("C" & m).Value = l5value And masterData.Range("BL" & m).Value = l6value Then

                            timeInfo.Add masterData.Range("AL" & m).Value

                    End If

                Next m

               

                leveldataRange = "A" & CStr(calcSheetrowcount) & ":" & "F" & CStr(calcSheetrowcount)

                calcRange = "G" & CStr(calcSheetrowcount) & ":" & "R" & CStr(calcSheetrowcount)

                

                'calculate statistics

                arrayDataInfoNoOutliers = removeOutliers(timeInfo)

                arrayDataInfo = toArray(timeInfo)

               

                averageinfo(0) = calculateAverage(arrayDataInfoNoOutliers)

                averageinfo(1) = calculateAverage(arrayDataInfo)

                stddevinfo(0) = calculateStandardDeviation(arrayDataInfoNoOutliers)

                stddevinfo(1) = calculateStandardDeviation(arrayDataInfo)

                confidenceIntervalinfo(0) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(0)

                confidenceIntervalinfo(1) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(1)

                confidenceIntervalinfo(2) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(0)

                confidenceIntervalinfo(3) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(1)

                countinfo(0) = calculateCount(arrayDataInfoNoOutliers)

                countinfo(1) = calculateCount(arrayDataInfo)

                totalcountpercent(0) = countinfo(0) / (masterDatarowcount - 1)

                totalcountpercent(1) = countinfo(1) / (masterDatarowcount - 1)

               

                

                'add calculations

                calc(1) = averageinfo(0)

                calc(2) = averageinfo(1)

                calc(3) = stddevinfo(0)

                calc(4) = stddevinfo(1)

                calc(5) = confidenceIntervalinfo(0)

                calc(6) = confidenceIntervalinfo(1)

                calc(7) = confidenceIntervalinfo(2)

                calc(8) = confidenceIntervalinfo(3)

                calc(9) = countinfo(0)

                calc(10) = countinfo(1)

                calc(11) = totalcountpercent(0)

                calc(12) = totalcountpercent(1)

               

                

                'output

                calcSheet.Range(leveldataRange).Value = leveldata

                calcSheet.Range(calcRange).Value = calc

               

                Set timeInfo = New Collection

               

                calcSheetrowcount = calcSheetrowcount + 1

                  

            Next l

   

    

    'clear objects

    Set masterData = Nothing

    Set levels = Nothing

    Set calcSheet = Nothing

   

    MsgBox "high low Complete"

   

    Application.ScreenUpdating = True

 

 

Exit Sub

errorhandler:

    MsgBox Err.Description

    End

 

End Sub

 

 

 

 

Public Sub driver_level_high_low()

 

On Error GoTo errorhandler

 

    Dim masterData As Worksheet, levels As Worksheet, calcSheet As Worksheet 'worksheets

   

    Dim masterDatarowcount As Long, calcSheetrowcount As Long, levelrowcount As Long, dayrowcount As Long 'row count

   

    Dim l1value As String, l2value As String, l3value As String, l4value As String, l5value As String, l6value As String 'level information

   

    Dim l As Long, m As Long, s As Long 'counters

   

    Dim averageinfo(0 To 1) As Double

    Dim stddevinfo(0 To 1) As Double

    Dim confidenceIntervalinfo(0 To 3) As Double

    Dim countinfo(0 To 1) As Long

   

    Dim arrayDataInfo As Variant

    Dim arrayDataInfoNoOutliers As Variant

    Dim timeInfo As New Collection

   

    Dim leveldata(1 To 5) As Variant

    Dim calc(1 To 12) As Double

   

    Dim calcRange As String

    Dim leveldataRange As String

    Dim totalcountpercent(0 To 1) As Double

   

    

    Set masterData = ThisWorkbook.Sheets("Master Data")

    Set levels = ThisWorkbook.Sheets("High_Low_levels_2")

    Set calcSheet = ThisWorkbook.Sheets("Calc_Time_High_Low_2")

   

 

    'clear old data

    Application.ScreenUpdating = False

    calcSheet.Rows("2:100000").Delete Shift:=xlUp

   

    masterDatarowcount = masterData.Cells(masterData.Rows.Count, "D").End(xlUp).Row

    calcSheetrowcount = 2

    levelrowcount = levels.Cells(levels.Rows.Count, "A").End(xlUp).Row

 

   

            For l = 2 To levelrowcount

                l1value = levels.Range("A" & l).Value

                l2value = levels.Range("B" & l).Value

                l3value = levels.Range("C" & l).Value

                l4value = levels.Range("D" & l).Value

                l5value = levels.Range("E" & l).Value

                'l6value = levels.Range("F" & l).Value

               

                leveldata(1) = l1value

                leveldata(2) = l2value

                leveldata(3) = l3value

                leveldata(4) = l4value

                leveldata(5) = l5value

                'leveldata(6) = l6value

               

                For m = 2 To masterDatarowcount

                    If masterData.Range("U" & m).Value = l1value And masterData.Range("V" & m).Value = l2value And _

                        masterData.Range("W" & m).Value = l3value And masterData.Range("X" & m).Value = l4value And _

                        masterData.Range("BL" & m).Value = l5value Then

                            timeInfo.Add masterData.Range("AL" & m).Value

                    End If

                Next m

               

                leveldataRange = "A" & CStr(calcSheetrowcount) & ":" & "E" & CStr(calcSheetrowcount)

                calcRange = "F" & CStr(calcSheetrowcount) & ":" & "Q" & CStr(calcSheetrowcount)

               

                'calculate statistics

                arrayDataInfoNoOutliers = removeOutliers(timeInfo)

                arrayDataInfo = toArray(timeInfo)

               

                averageinfo(0) = calculateAverage(arrayDataInfoNoOutliers)

                averageinfo(1) = calculateAverage(arrayDataInfo)

                stddevinfo(0) = calculateStandardDeviation(arrayDataInfoNoOutliers)

                stddevinfo(1) = calculateStandardDeviation(arrayDataInfo)

                confidenceIntervalinfo(0) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(0)

                confidenceIntervalinfo(1) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(1)

                confidenceIntervalinfo(2) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(0)

                confidenceIntervalinfo(3) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(1)

                countinfo(0) = calculateCount(arrayDataInfoNoOutliers)

                countinfo(1) = calculateCount(arrayDataInfo)

                totalcountpercent(0) = countinfo(0) / (masterDatarowcount - 1)

                totalcountpercent(1) = countinfo(1) / (masterDatarowcount - 1)

               

                

                'add calculations

                calc(1) = averageinfo(0)

                calc(2) = averageinfo(1)

                calc(3) = stddevinfo(0)

                calc(4) = stddevinfo(1)

                calc(5) = confidenceIntervalinfo(0)

                calc(6) = confidenceIntervalinfo(1)

                calc(7) = confidenceIntervalinfo(2)

                calc(8) = confidenceIntervalinfo(3)

                calc(9) = countinfo(0)

                calc(10) = countinfo(1)

                calc(11) = totalcountpercent(0)

                calc(12) = totalcountpercent(1)

               

                

                'output

                calcSheet.Range(leveldataRange).Value = leveldata

                calcSheet.Range(calcRange).Value = calc

               

                Set timeInfo = New Collection

               

                calcSheetrowcount = calcSheetrowcount + 1

                  

            Next l

   

    

    'clear objects

    Set masterData = Nothing

    Set levels = Nothing

    Set calcSheet = Nothing

   

    MsgBox "high low 2 Complete"

   

    Application.ScreenUpdating = True

 

 

Exit Sub

errorhandler:

    MsgBox Err.Description

    End

 

End Sub

 

 

Public Sub driver_high_low()

 

On Error GoTo errorhandler

 

    Dim masterData As Worksheet, levels As Worksheet, calcSheet As Worksheet 'worksheets

   

    Dim masterDatarowcount As Long, calcSheetrowcount As Long, levelrowcount As Long, dayrowcount As Long 'row count

   

    Dim l1value As String, l2value As String, l3value As String, l4value As String, l5value As String, l6value As String 'level information

   

    Dim l As Long, m As Long, s As Long 'counters

   

    Dim averageinfo(0 To 1) As Double

    Dim stddevinfo(0 To 1) As Double

    Dim confidenceIntervalinfo(0 To 3) As Double

    Dim countinfo(0 To 1) As Long

   

    Dim arrayDataInfo As Variant

    Dim arrayDataInfoNoOutliers As Variant

    Dim timeInfo As New Collection

   

    Dim leveldata(1 To 5) As Variant

    Dim calc(1 To 12) As Double

   

    Dim calcRange As String

    Dim leveldataRange As String

    Dim totalcountpercent(0 To 1) As Double

   

    

    Set masterData = ThisWorkbook.Sheets("Master Data")

    Set levels = ThisWorkbook.Sheets("high_low")

    Set calcSheet = ThisWorkbook.Sheets("Calc_High_Low")

   

 

    'clear old data

    Application.ScreenUpdating = False

    calcSheet.Rows("2:100000").Delete Shift:=xlUp

   

    masterDatarowcount = masterData.Cells(masterData.Rows.Count, "D").End(xlUp).Row

    calcSheetrowcount = 2

    levelrowcount = levels.Cells(levels.Rows.Count, "A").End(xlUp).Row

 

   

            For l = 2 To levelrowcount

                l1value = levels.Range("A" & l).Value

                'l2value = levels.Range("B" & l).Value

                'l3value = levels.Range("C" & l).Value

                'l4value = levels.Range("D" & l).Value

                'l5value = levels.Range("E" & l).Value

                'l6value = levels.Range("F" & l).Value

               

                'leveldata(1) = l1value

                'leveldata(2) = l2value

                'leveldata(3) = l3value

                'leveldata(4) = l4value

                'leveldata(5) = l5value

                'leveldata(6) = l6value

               

                For m = 2 To masterDatarowcount

                    If masterData.Range("BN" & m).Value = l1value Then

                            timeInfo.Add masterData.Range("AL" & m).Value

                    End If

                Next m

               

                leveldataRange = "A" & CStr(calcSheetrowcount) & ":" & "A" & CStr(calcSheetrowcount)

                calcRange = "B" & CStr(calcSheetrowcount) & ":" & "M" & CStr(calcSheetrowcount)

               

                'calculate statistics

                arrayDataInfoNoOutliers = removeOutliers(timeInfo)

                arrayDataInfo = toArray(timeInfo)

               

                averageinfo(0) = calculateAverage(arrayDataInfoNoOutliers)

                averageinfo(1) = calculateAverage(arrayDataInfo)

                stddevinfo(0) = calculateStandardDeviation(arrayDataInfoNoOutliers)

                stddevinfo(1) = calculateStandardDeviation(arrayDataInfo)

                confidenceIntervalinfo(0) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(0)

                confidenceIntervalinfo(1) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(1)

                confidenceIntervalinfo(2) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(0)

                confidenceIntervalinfo(3) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(1)

                countinfo(0) = calculateCount(arrayDataInfoNoOutliers)

                countinfo(1) = calculateCount(arrayDataInfo)

                totalcountpercent(0) = countinfo(0) / (masterDatarowcount - 1)

                totalcountpercent(1) = countinfo(1) / (masterDatarowcount - 1)

                

                

                'add calculations

                calc(1) = averageinfo(0)

                calc(2) = averageinfo(1)

                calc(3) = stddevinfo(0)

                calc(4) = stddevinfo(1)

                calc(5) = confidenceIntervalinfo(0)

                calc(6) = confidenceIntervalinfo(1)

                calc(7) = confidenceIntervalinfo(2)

                calc(8) = confidenceIntervalinfo(3)

                calc(9) = countinfo(0)

                calc(10) = countinfo(1)

                calc(11) = totalcountpercent(0)

                calc(12) = totalcountpercent(1)

               

                

                'output

                calcSheet.Range(leveldataRange).Value = l1value 'leveldata

                calcSheet.Range(calcRange).Value = calc

               

                Set timeInfo = New Collection

               

                calcSheetrowcount = calcSheetrowcount + 1

                  

            Next l

   

    

    'clear objects

    Set masterData = Nothing

    Set levels = Nothing

    Set calcSheet = Nothing

   

    MsgBox "high low Complete"

   

    Application.ScreenUpdating = True

 

 

Exit Sub

errorhandler:

    MsgBox Err.Description

    End

 

End Sub

 

 

Public Sub driver_main()

 

On Error GoTo errorhandler

 

    Dim masterData As Worksheet, levels As Worksheet, calcSheet As Worksheet 'worksheets

   

    Dim masterDatarowcount As Long, calcSheetrowcount As Long, levelrowcount As Long, dayrowcount As Long 'row count

   

    Dim l1value As String, l2value As String, l3value As String, l4value As String, l5value As String, l6value As String 'level information

   

    Dim l As Long, m As Long, s As Long 'counters

   

    Dim averageinfo(0 To 1) As Double

    Dim stddevinfo(0 To 1) As Double

    Dim confidenceIntervalinfo(0 To 3) As Double

    Dim countinfo(0 To 1) As Long

   

    Dim arrayDataInfo As Variant

    Dim arrayDataInfoNoOutliers As Variant

    Dim timeInfo As New Collection

   

    Dim leveldata(1 To 5) As Variant

    Dim calc(1 To 12) As Double

   

    Dim calcRange As String

    Dim leveldataRange As String

    Dim totalcountpercent(0 To 1) As Double

   

    

    Set masterData = ThisWorkbook.Sheets("Master Data")

    Set levels = ThisWorkbook.Sheets("Drivers")

    Set calcSheet = ThisWorkbook.Sheets("Calc_Drivers")

   

 

    'clear old data

    Application.ScreenUpdating = False

    calcSheet.Rows("2:100000").Delete Shift:=xlUp

   

    masterDatarowcount = masterData.Cells(masterData.Rows.Count, "D").End(xlUp).Row

    calcSheetrowcount = 2

    levelrowcount = levels.Cells(levels.Rows.Count, "A").End(xlUp).Row

 

   

            For l = 2 To levelrowcount

                l1value = levels.Range("A" & l).Value

                'l2value = levels.Range("B" & l).Value

                'l3value = levels.Range("C" & l).Value

                'l4value = levels.Range("D" & l).Value

                'l5value = levels.Range("E" & l).Value

                'l6value = levels.Range("F" & l).Value

               

                'leveldata(1) = l1value

                'leveldata(2) = l2value

                'leveldata(3) = l3value

                'leveldata(4) = l4value

                'leveldata(5) = l5value

                'leveldata(6) = l6value

               

                For m = 2 To masterDatarowcount

                    If masterData.Range("BM" & m).Value = l1value Then

                            timeInfo.Add masterData.Range("AL" & m).Value

                    End If

                Next m

               

                leveldataRange = "A" & CStr(calcSheetrowcount) & ":" & "A" & CStr(calcSheetrowcount)

                calcRange = "B" & CStr(calcSheetrowcount) & ":" & "M" & CStr(calcSheetrowcount)

               

                'calculate statistics

                arrayDataInfoNoOutliers = removeOutliers(timeInfo)

                arrayDataInfo = toArray(timeInfo)

               

                averageinfo(0) = calculateAverage(arrayDataInfoNoOutliers)

                averageinfo(1) = calculateAverage(arrayDataInfo)

                stddevinfo(0) = calculateStandardDeviation(arrayDataInfoNoOutliers)

                stddevinfo(1) = calculateStandardDeviation(arrayDataInfo)

                confidenceIntervalinfo(0) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(0)

                confidenceIntervalinfo(1) = calculateConfidenceInterval(arrayDataInfoNoOutliers, averageinfo(0), stddevinfo(0))(1)

                confidenceIntervalinfo(2) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(0)

                confidenceIntervalinfo(3) = calculateConfidenceInterval(arrayDataInfo, averageinfo(1), stddevinfo(1))(1)

                countinfo(0) = calculateCount(arrayDataInfoNoOutliers)

                countinfo(1) = calculateCount(arrayDataInfo)

                totalcountpercent(0) = countinfo(0) / (masterDatarowcount - 1)

                totalcountpercent(1) = countinfo(1) / (masterDatarowcount - 1)

               

                

                'add calculations

                calc(1) = averageinfo(0)

                calc(2) = averageinfo(1)

                calc(3) = stddevinfo(0)

                calc(4) = stddevinfo(1)

                calc(5) = confidenceIntervalinfo(0)

                calc(6) = confidenceIntervalinfo(1)

                calc(7) = confidenceIntervalinfo(2)

                calc(8) = confidenceIntervalinfo(3)

                calc(9) = countinfo(0)

                calc(10) = countinfo(1)

                calc(11) = totalcountpercent(0)

                calc(12) = totalcountpercent(1)

               

                

                'output

                calcSheet.Range(leveldataRange).Value = l1value 'leveldata

                calcSheet.Range(calcRange).Value = calc

               

                Set timeInfo = New Collection

               

                calcSheetrowcount = calcSheetrowcount + 1

                  

            Next l

   

    

    'clear objects

    Set masterData = Nothing

    Set levels = Nothing

    Set calcSheet = Nothing

   

    MsgBox "drivers Complete"

   

    Application.ScreenUpdating = True

 

 

Exit Sub

errorhandler:

    MsgBox Err.Description

    End

 

End Sub

 




