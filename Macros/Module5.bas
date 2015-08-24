Attribute VB_Name = "Module5"
Sub Tier2TransposeMacro()
Attribute Tier2TransposeMacro.VB_Description = "This Macro is responsible for transposing the data in the table that is dowloaded from Caspio."
Attribute Tier2TransposeMacro.VB_ProcData.VB_Invoke_Func = "i\n14"
'
' Tier2TransposeMacro Macro
' This Macro is responsible for transposing the data in the table that is downloaded from Caspio.
'

'
    Sheets("Tier2_Quarterly_Data").Select
    ActiveSheet.UsedRange.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet3").Select
    Sheets("Sheet3").Name = "Tier2_Actual"
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Selection.Columns.AutoFit
    Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Selection.Cut
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "COMPANY NAME "
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "CONFIDENTIAL"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "AB 2398 Quarterly Report - Tier 2 Manufacturer"
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Selection.Cut
    Rows("35:35").Select
    Selection.Insert Shift:=xlDown
    ActiveWindow.SmallScroll Down:=-45
    Rows("4:4").Select
    Selection.Cut
    Rows("35:35").Select
    Selection.Insert Shift:=xlDown
    ActiveWindow.SmallScroll Down:=-42
    Rows("5:5").Select
    Selection.Delete Shift:=xlUp
    Rows("6:6").Select
    Selection.Delete Shift:=xlUp
    Rows("7:7").Select
    Selection.Delete Shift:=xlUp
    Rows("8:8").Select
    Selection.Delete Shift:=xlUp
    Rows("9:9").Select
    Selection.Delete Shift:=xlUp
    Rows("10:10").Select
    Selection.Delete Shift:=xlUp
    Rows("11:11").Select
    Selection.Delete Shift:=xlUp
    Rows("12:12").Select
    Selection.Delete Shift:=xlUp
    Rows("13:13").Select
    Selection.Delete Shift:=xlUp
    Rows("14:14").Select
    Selection.Delete Shift:=xlUp
    Rows("15:15").Select
    Selection.Delete Shift:=xlUp
    Rows("16:16").Select
    Selection.Delete Shift:=xlUp
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A4").Select
    ActiveCell.FormulaR1C1 = _
        "If Located in CA Number of Full Time Equivalent (FTE) Employees working on PCC Products"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = _
        "Number of FTE CA Employees at end of this quarter using PCC carpet?"
    Rows("6:6").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A6").Select
    ActiveCell.FormulaR1C1 = _
        "Type 1, Non-Nylon PC Carpet pounds purchased by you this quarter"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = _
        "Type 1 pounds directly purchased by you from a QUALIFIED Processor of CA Waste Carpet this quarter?"
    Rows("8:8").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "Please supply confirmation letter from supplier"
    Rows("9:9").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A9").Select
    ActiveCell.FormulaR1C1 = _
        "Type 1, Non-Nylon Processed CA PC Carpet pounds directly purchased by YOU by FIBER type"
    Range("A10").Select
    ActiveCell.FormulaR1C1 = "Polypropylene"
    Range("A11").Select
    ActiveCell.FormulaR1C1 = "PET"
    Range("A12").Select
    ActiveCell.FormulaR1C1 = "Other including mixed non-nylon fibers"
    Rows("13:13").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A13").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("14:14").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A14").Select
    ActiveCell.FormulaR1C1 = "Line 13 must equal Line 7"
    Rows("15:15").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A15").Select
    ActiveCell.FormulaR1C1 = _
        "Accounting for total processed Type 1 PC Carpet Inputs & Beginning Inventory this quarter"
    Range("A16").Select
    ActiveCell.FormulaR1C1 = _
        "Beginning Inventory of Type 1 Non-Nylon processed PC Carpet from CA at start of quarter (should equal prior quarter ending inventory)."
    Range("A17").Select
    ActiveCell.FormulaR1C1 = _
        "Type 1 Non-Nylon Processed PC Carpet received/purchased (Row 7)"
    Rows("18:18").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A18").Select
    ActiveCell.FormulaR1C1 = "Total Material Available for Current Quarter "
    Rows("19:19").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A19").Select
    ActiveCell.FormulaR1C1 = _
        "Accounting for total PC Carpet Outputs & Ending Inventory"
    Range("A20").Select
    ActiveCell.FormulaR1C1 = _
        "Type 1 Non-Nylon Processed PC Carpet SOLD & SHIPPED this quarter? [SEE NOTE 1]"
    Rows("21:21").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A21").Select
    ActiveCell.FormulaR1C1 = _
        "Output and other destinations of Non-Nylon Type 1 materials internally processed this quarter"
    Rows("21:21").Select
    Rows("21:21").Select
    Rows("22:22").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A22").Select
    ActiveCell.FormulaR1C1 = "Tier 2 Non-Nylon Products SOLD & SHIPPED in Quarter"
    Range("A26").Select
    ActiveCell.FormulaR1C1 = _
        "Total Requested ($) Tier 2 Non-Nylon Output, $0.12/lb."
    Range("A23").Select
    ActiveCell.FormulaR1C1 = ""
    Range("A24").Select
    ActiveCell.FormulaR1C1 = ""
    Range("A25").Select
    ActiveCell.FormulaR1C1 = ""
    Rows("26:26").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A26").Select
    ActiveCell.FormulaR1C1 = "Calculations for funding"
    Rows("28:28").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "2"
    Selection.AutoFill Destination:=Range("A2:A27"), Type:=xlFillSeries
    Range("A2:A27").Select
    Columns("A:A").ColumnWidth = 4
    Columns("B:B").ColumnWidth = 44.78
    Columns("B:B").ColumnWidth = 52.78
    
    '********************************************************************'
    'To find the first empty column on the table to insert the total'
    'find last used cell on the row to the right
    Range("A1").Select
    ActiveCell.End(xlToRight).Select
    'move one cell to the right from the last used cell
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "Total"
    
    
    
    '*********************************************************************'
    'Select the table to insert the borders'
    Range("A1:F27").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B1:E3").Select
    Selection.Font.Bold = True
    Range("B2").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:E27").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("B:B").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B4:E4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("A4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B9:E9").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("A9:E9").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B15:E15").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("A15:E15").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B19:E19").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.Font.Bold = True
    Range("A19:E19").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B21:E21").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Font.Bold = True
    Range("A21:E21").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B15:E15").Select
    Selection.Font.Bold = True
    Range("B9:E9").Select
    Selection.Font.Bold = True
    Range("B4:E4").Select
    Selection.Font.Bold = True
    Range("B26:E26").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.Font.Bold = True
    Range("A26:E26").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B20").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("B27").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("B27").Select
    Selection.End(xlToRight).Select
    Rows("27:27").Select
    Selection.Font.Bold = True
    Range("B22").Select
    Selection.Font.Bold = True
    Range("B16").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B7").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("B:B").ColumnWidth = 56
    Rows("16:16").Select
    Selection.Rows.AutoFit
    Range("B14").Select
    Columns("E:E").EntireColumn.AutoFit
    Range("C18").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-2]C:R[-1]C)"
    Range("C18").Select
    Selection.AutoFill Destination:=Range("C18:E18"), Type:=xlFillDefault
    Range("C18:E18").Select
    Range("C13").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
    Range("C13").Select
    Selection.AutoFill Destination:=Range("C13:E13"), Type:=xlFillDefault
    Range("C13:E13").Select
End Sub
