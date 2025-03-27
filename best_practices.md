# VBA

[Optimize VBA code for performance](https://vbacompiler.com/optimize-vba-code/)

## 1. Turn off “Automatic Calculation” mode and enable “Manual Calculation” mode

 

The Application.Calculation property allows you to control Excel’s calculation mode from your VBA code.

Automatic (xlCalculationAutomatic) mode is the default value of Application.Calculation property and it means that Excel controls the calculation and decides when to trigger the calculation of the workbook (e.g. when a new formula is entered, a cell value is changed, or the value of the Range object is changed from VBA code).

Manual (xlCalculationManual) mode means Excel waits for the user action (or action from your VBA code) to explicitly begin calculation.
Since recalculating a workbook takes time and resources, then to increase the VBA code calculation performance it is better to set the Application.Calculation property to manual mode.

Turn off the automatic calculation mode while your VBA code is running, and then set the mode back when it’s done to increase the VBA performance.

Example:
`
Dim savedCalcMode As XlCalculation
savedCalcMode = Application.Calculation
Application.Calculation = xlCalculationManual
< ... your code here ... >
Application.Calculation = savedCalcMode
`
 

## 2. Turn off “Screen Updating”

The Excel Application.ScreenUpdating property controls the re-drawing of the Excel screen’s visible parts. The Excel spends some resources to draw the screen during re-calculation. You will get noticeable performance improvements by switching off the Application.ScreenUpdating property by setting it to False.

Example:
`
Dim savedScreenUpdating as Boolean
savedScreenUpdating = Application.ScreenUpdating
Application.ScreenUpdating = False
< ... your code here ... >
Application.ScreenUpdating = savedScreenUpdating
`
 

## 3. Disable Application.EnableEvents

In most cases you do not want to allow the Excel to process events while the VBA code is running because it slows calculation.

Example:
`
Dim savedEnableEvents as Boolean
savedEnableEvents = Application.EnableEvents
Application.EnableEvents = False
< ... your code here ... >
Application.EnableEvents = savedEnableEvents
 `

## 4. Turn off ActiveSheet.DisplayPageBreaks

 

Another way to speed up your VBA code is by disabling the ActiveSheet.DisplayPageBreaks property.

Example:
`
Dim savedPageBrakes as Boolean
savedPageBrakes = ActiveSheet. DisplayPageBreaks
ActiveSheet. DisplayPageBreaks = False
< ... your code here ... >
ActiveSheet. DisplayPageBreaks = savedPageBrakes
` 

## 5. Disable animation in Excel by “Application.EnableAnimations = False”

 

Animations are enabled for user actions and representations, beautifying Excel. Disabling these animations allows for the improvement of VBA performance.

Example:
`
Dim savedEnableAnimations as Boolean
savedEnableAnimations = Application.EnableAnimations
Application.EnableAnimations = False
< ... your code here ... >
Application.EnableAnimations = savedEnableAnimations
 `

## 6. Turn off status bar by “Application.DisplayStatusBar = False”

 

Disabling the status bar frees up resources for VBA calculation and improves VBA calculation performance.

Example:
`
Dim savedStatusBar as Boolean
savedStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = False
< ... your code here ... >
Application.DisplayStatusBar = savedStatusBar
 `

## 7. Turn off “Print Communication”

 

Disabling Application.PrintCommunication property also speeds up VBA execution.

Example:
`
Dim savedPrintCommunication as Boolean
savedPrintCommunication = Application.PrintCommunication
Application.PrintCommunication = False
< ... your code here ... >
Application.PrintCommunication = savedPrintCommunication
 `

## 8. Use OptimizedMode() procedure to enable optimization

 

The best way to use the actions from the tips above is to combine them all into one procedure which allows you to enable an optimized VBA calculation mode with a single call before executing your VBA code and disabling it all with one call after that.

 

Here you may find the full OptimizedMode procedure definition.
`
Public Sub OptimizedMode(ByVal enable As Boolean)
     Application.EnableEvents = Not enable
     Application.Calculation = IIf(enable, xlCalculationManual, xlCalculationAutomatic)
     Application.ScreenUpdating = Not enable
     Application.EnableAnimations = Not enable
     Application.DisplayStatusBar = Not enable
     Application.PrintCommunication = Not enable
End Sub
`

Usage Example:

OptimizedMode True
< ... your code here ... >
OptimizedMode False
 

## 9. Avoid selecting and activating methods

 

Select and Activate methods are slow methods and add a time overhead during calculation.

For example:
Instead of using ‘Select’ method and ‘Selection’ object, like:

`
ActiveSheet.Select
Selection.Range("A1").Value2 = "text"
`
Use direct access to the object without unnecessary selection:

`
ActiveSheet.Range("A1").Value2 = "text"
`

## 10. Avoid recording macro
VBA code generated from macro recording has low performance efficiency because it uses a lot of Select and Selection methods.

 

## 11. Avoid unnecessary Copy/Paste operations
Copy/Paste operations may seem appropriate for many algorithm situations, but it is very ineffective for VBA code performance.

Instead of the code:
`
Worksheets("Sheet1").Range("A1:H10").Copy
Worksheets("Sheet2").Range("A1").PasteSpecial
`
Use the following code:
`
Sheet2.Range("A1:H10").Value2 = Sheet1.Range("A1:H10").Value2
` 

## 12. Use “ForEach” loop instead of “indexed For” loop for collections

 

When a collection object is traversed in an indexed For loop, each item in the collection is accessed through an index operation, which is much slower than traversing each object in the collection one-by-one in the ‘For Each’ loop.

Instead of code like:
`
For i=1 to Collection.Count
    Collection.Item(i).Value = value
Next 
`
Use following approach:
`
For Each obj in Collection
    Obj.Value = value
Next
`
 

## 13. Use “With … End” statement

Use the With…End statement instead of full qualified access to the object’s methods or data.

Instead of:
`
ThisWorkbook.Worksheets("Sheet1").Range("A1").Font.Bold = True
ThisWorkbook.Worksheets("Sheet1").Range("A1").Value2 = 100
`

Use the following code:
`
With ThisWorkbook.Worksheets("Sheet1").Range("A1")
    .Font.Bold = True
    .Value = 100
End With
`

## 14. Use simple objects rather than compound objects

 

For the same reason as in the previous advice, you can reduce the time overhead for accessing the qualified object’s method by assigning a complex qualified name to a separate object and use that object for accessing the object’s methods.

Instead of:

`
ThisWorkbook.Worksheets("Sheet1").Range("A1").Font.Bold = True
ThisWorkbook.Worksheets("Sheet1").Range("A1").Value2 = 100
`
Use the following code:

`
Dim rng as Range
Set rng = ThisWorkbook.Worksheets("Sheet1").Range("A1")
Rng.Font.Bold = True
Rng.Value2 = 100
`

## 15. Use vbNullString instead of an empty string
Technically vbNullString is not a string – it is a constant declared as a Null pointer. When you use vbNullString instead of empty string (“”) you avoid some extra memory allocations which give time overheads.

Instead of assignment like this:

`
 Dim Str as String
 Str = ""
`

Use this:
`
 Dim Str as String
 Str = vbNullString
` 

## 16. Use “Early Binding” instead of “Late Binding”

Binding is a process of coupling the object variable with its content. After the binding, the variable represents the object and allows it to get access to the methods and data of the object.
For detailed explanation of the difference between ‘Early binding’ and ‘Late Binding’ please read the following articles:
Early vs. Late Binding
Using early binding and late binding in Automation

Example of “Late binding”:

`
Dim obj as Object
Set obj = CreateObject("Excel.Application")
`

Example of “Early Binding”:

`
Dim xlApp as Excel.Application
Set xlApp = New Excel.Application
`

## 17. Avoid using Variant type


Variant type is a dynamic type – that means Variant variables can accept any other type of data and store it. Usage of the Variant type may be comfortable and time saving for programming time, but it sacrifices time overhead during run-time (when your code is working) because conversion from exact type to Variant and from Variant to any other type is time consuming.
For example: If the data has a ‘String’ type it should be stored into the variable which is declared with a String type but not into a Variant type variable, also it makes sense for all other types.

Also avoid declaration like:
`
 Dim v
 `
The variable ‘v’ in code line above is declared as Variant because declaration of type is missed.
`
 Dim a, b, c as Long
 `
In this code example only the ‘c’ variable is declared as Long, but ‘a’ and ‘b’ are declared as Variant.
To declare all variables in one line you need to point out the exact type for each variable declaration, like so:
`
 Dim a as Long, b as Long, c as Long
`
## 18. Avoid declaration of method parameters without type

Instead of:
`
Private Sub Proc1(a,b,c)
`
Use exact type declaration for each parameter:

`
Private Sub Proc1(ByVal a as Long, ByVal b as String, ByVal c as Integer)
 `

## 19. Use ‘Option Explicit’

If you miss the variable declaration it will be assumed by VBA as variable with Variant type.

Following to the previous tip it is better to avoid Variant typed declarations if you want to maximize your VBA code performance.

By default, the VBA do not warn you that variable is not declared. To avoid this situation, use the ‘Option Explicit’ statement in the first line of each VBA module in your VBA project.

With the ‘Option Explicit’ statement, VBA will warn about any undeclared variable.

Optimize VBA code tip - Option Explicit directive allows to detect undeclared variables

 

## 20. Avoid using Object type, use specific object type declarations

When you use the Object type, VBA does not know what exact type object it contains and will resolve this during run-time (when your code is working) which will take additional time. To avoid this overhead time, you need to use the exact type of variable for declaration.

Instead of code:

`
Dim obj as Object
Set obj = New Collection
 `

Use following code:

`
Dim oCollection as Collection
Set oCollection = New Collection
 `

## 21. Use “Const” for constant declarations

 

Declare constant instead of variable if you need to have a constant.

Use constant declaration:
`
Const pi as Double = 3.14159265359
`
Instead of following:

`
Dim pi as Double
Pi = 3.14159265359
 `

## 22. Use ByVal modifiers for parameters

 

When you do not need to return value from the method parameter then explicitly declare this parameter as ByVal.

Use this:
`
Private Sub Proc1(ByVal a as Long)
 `

Instead of:

`
Private Sub Proc1(a as Long)
 `

## 23. Move out unnecessary actions from a loop body

 

Each unnecessary action inside a loop will be repeated and it will accumulate overhead time for algorithm. By removing such pieces of code from loop body you optimize VBA code.

For example, instead of the following code:

`
Dim rng as Range
For Each rng in Worksheets("Sheet1").Range("A1:AA100")
    Dim result as Range
    Set result = Worksheets("Sheet2").Range("A1")
    result.Value2 = result.Value2 + rng.Value2
Next
`

Use this code:

`
Dim result As Double
Dim rng As Range
For Each rng In Worksheets("Sheet1").Range("A1:AA100")
    result = result + CDbl(rng.Value2)
Next
Worksheets("Sheet3").Range("A1").Value2 = result
` 

## 24. Use String versions of built-in functions instead of the Variant version

If you are working with VBA string functions, it is better to use String typed version functions – which have a ‘$’ dollar sign in the suffix instead of the Variant typed version of the same functions.

Instead of these functions:

`
Chr(), ChrB(), ChrW(), Error(), Format(), Hex(), LCase(), Mid(), MidB(), Left(), LeftB(), 
LTrim(), Oct(), RightB(), Right(), RTrim(), Space(), Str(), String(),Trim(), UCase() 
`
Use these string version functions:
`
Chr$(), ChrB$(), ChrW$(), Error$(), Format$(), Hex$(), LCase$(), Mid$(), MidB$(), Left$(),
LeftB$(), LTrim$(), Oct$(), RightB$(), Right$(), RTrim$(), Space$(), Str$(), String$(),
Trim$(), UCase$()
`
 

## 25. Minimize data exchange between Worksheet and VBA code

 

Use an array to collect big range data and traverse through the array rather than a Worksheet range, cell by cell. This optimize VBA code tip allows to improve the performance more than 10 times.

Instead of this approach:

Function CalcSlow() As Double
    Dim cell As Range
    For Each cell In Sheet1.UsedRange
        If IsNumeric(cell.Value) Then
            CalcSlow = CalcSlow + cell.Value
        End If
    Next
End Function
Use this approach:

'the CalcFast() works more than x10 times faster of CalcSlow()
Function CalcFast() As Double
    Dim arrCells() As Variant
    Dim val As Variant
    arrCells = Sheet1.UsedRange.Value2
    For Each val In arrCells
        If IsNumeric(val) Then
            CalcFast = CalcFast + CDbl(val)
        End If
    Next
End Function
