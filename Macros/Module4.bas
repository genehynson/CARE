Attribute VB_Name = "Module4"
Sub ForecastMacro()
Attribute ForecastMacro.VB_Description = "This macro is responsible for formatting the Tier 1 Forecast report."
Attribute ForecastMacro.VB_ProcData.VB_Invoke_Func = "f\n14"
'
' ForecastMacro Macro
' This macro is responsible for formatting the Tier 1 Forecast report.
'
' Keyboard Shortcut: Ctrl+f
'
    Application.Run "PERSONAL.XLSB!Tier1ActualMacro"
    Sheets("Tier1_Forecast").Select
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Rows("3:3").Select
    Selection.Cut
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "CONFIDENTIAL"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "AB 2398 Monthly Rolling Forecast"
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A5").Select
    ActiveCell.FormulaR1C1 = _
        "Number of CA FTE Employees at the beginning of this quarter"
    Range("A6").Select
    Columns("A:A").EntireColumn.AutoFit
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "Number of FTE CA Jobs lost this quarter"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "Number of FTE CA Jobs gained this quarter"
    Range("A8").Select
    ActiveCell.FormulaR1C1 = "Number of FTE CA Employees at end of this quarter"
    Rows("9:9").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A10").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds directly collected by you from California for this quarter"
    Range("A11").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds directly collected by you from California for this quarter"
    Range("A11").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds directly collected by you from OUTSIDE California for this quarter"
    Rows("12:12").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A12").Select
    ActiveCell.FormulaR1C1 = "TOTAL Post-consumer carpet pounds"
    Rows("13:13").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("14:14").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A14").Select
    ActiveCell.FormulaR1C1 = "Nylon 6"
    Rows("15:15").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "Nylon6,6"
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "Nylon 6,6"
    Rows("16:16").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A16").Select
    ActiveCell.FormulaR1C1 = "Polypropylene"
    Rows("17:17").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A17").Select
    ActiveCell.FormulaR1C1 = "PET"
    Rows("18:18").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A18").Select
    ActiveCell.FormulaR1C1 = "Wool"
    Rows("19:19").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A19").Select
    ActiveCell.FormulaR1C1 = "Other/Mixed Fibers"
    Rows("20:20").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A20").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("21:21").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A21").Select
    ActiveCell.FormulaR1C1 = "Line 20 must equal line 10"
    Rows("22:22").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A23").Select
    ActiveCell.FormulaR1C1 = _
        "Beginning Inventory of Whole Carpet from CA at start of quarter (should equal prior quarter ending inventory)."
    Range("A24").Select
    ActiveCell.FormulaR1C1 = "Whole Carpet Collected from California (Row 10)"
    Range("A25").Select
    ActiveCell.FormulaR1C1 = "Whole Carpet from CA received from other collectors"
    Rows("26:26").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "T"
    Range("A26").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("27:27").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A28").Select
    ActiveCell.FormulaR1C1 = "Re-Used"
    Range("A29").Select
    ActiveCell.FormulaR1C1 = "Internally Used Whole Carpet this quarter"
    Range("A30").Select
    ActiveCell.FormulaR1C1 = _
        "Whole carpet shipped to US customers OUTSIDE California"
    Range("A31").Select
    ActiveCell.FormulaR1C1 = _
        "Whole carpet shipped to US customers OUTSIDE the United States"
    Range("A32").Select
    ActiveCell.FormulaR1C1 = "Whole carpet shipped to customers INSIDE California"
    Range("A33").Select
    ActiveCell.FormulaR1C1 = _
        "Non-carpet materials with value (i.e. carpet cushion)"
    Range("A34").Select
    ActiveCell.FormulaR1C1 = "WTE"
    Range("A35").Select
    ActiveCell.FormulaR1C1 = "Incinerated"
    Range("A36").Select
    ActiveCell.FormulaR1C1 = "Landfilled"
    Range("A37").Select
    ActiveCell.FormulaR1C1 = "Ending Inventory of Whole Carpet"
    Rows("38:38").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A38").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("39:39").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A39").Select
    ActiveCell.FormulaR1C1 = "Line 38 must equal line 26"
    Rows("40:40").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A41").Select
    ActiveCell.FormulaR1C1 = "Internally Used Whole Carpet"
    Range("A42").Select
    ActiveCell.FormulaR1C1 = "Processed"
    Range("A43").Select
    ActiveCell.FormulaR1C1 = "Landfilled"
    Range("A44").Select
    ActiveCell.FormulaR1C1 = "WTE"
    Range("A45").Select
    ActiveCell.FormulaR1C1 = "Incinerated"
    Rows("46:46").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A46").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("47:47").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A47").Select
    ActiveCell.FormulaR1C1 = "Line 46 must equal line 41"
    Rows("48:48").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A49").Select
    ActiveCell.FormulaR1C1 = _
        "Beginning Inventory of Processed Goods from prior quarter"
    Range("A50").Select
    ActiveCell.FormulaR1C1 = "Processed"
    Rows("51:51").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A51").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("52:52").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A52").Select
    ActiveCell.FormulaR1C1 = "Type 1 Outputs"
    Range("A53").Select
    ActiveCell.FormulaR1C1 = "Fiber"
    Range("A54").Select
    ActiveCell.FormulaR1C1 = "DePoly or Chemical Component"
    Range("A55").Select
    ActiveCell.FormulaR1C1 = "Shredded Carpet tile used for tile backing"
    Range("A56").Select
    ActiveCell.FormulaR1C1 = _
        "Number of Ash tests run this quarter (min 1 per 1M pounds)"
    Range("A57").Select
    ActiveCell.FormulaR1C1 = _
        "Average Ash Test Results over quarter for Type 1 pounds"
    Rows("58:58").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A58").Select
    ActiveCell.FormulaR1C1 = "Total Type 1 Ountput: SOLD & SHIPPED"
    Rows("59:59").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A59").Select
    ActiveCell.FormulaR1C1 = "Type 2 Outputs"
    Rows("60:60").Select
    Selection.Delete Shift:=xlUp
    Rows("60:60").Select
    Selection.Delete Shift:=xlUp
    Range("A60").Select
    ActiveCell.FormulaR1C1 = "Filler"
    Rows("61:61").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A61").Select
    ActiveCell.FormulaR1C1 = "Total Type 2 Output: SOLD & SHIPPED"
    Range("A62").Select
    ActiveCell.FormulaR1C1 = "CAAF"
    Range("A63").Select
    ActiveCell.FormulaR1C1 = "Cement Kiln feedstock"
    Range("A64").Select
    ActiveCell.FormulaR1C1 = "Carcass Sold"
    Range("A65").Select
    ActiveCell.FormulaR1C1 = "Landfilled"
    Range("A66").Select
    ActiveCell.FormulaR1C1 = "WTE"
    Range("A67").Select
    ActiveCell.FormulaR1C1 = "Incinerated"
    Range("A68").Select
    ActiveCell.FormulaR1C1 = "Ending Inventory Processed Goods this quarter"
    Rows("69:69").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A69").Select
    ActiveCell.FormulaR1C1 = "TOTAL Recycled Pounds This Quarter"
    Rows("70:70").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A70").Select
    ActiveCell.FormulaR1C1 = "Line 69 must equal line 51"
    Rows("71:71").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A72").Select
    ActiveCell.FormulaR1C1 = "Total_Payout_Adjustments"
    Range("A73").Select
    Rows("72:72").Select
    Selection.Delete Shift:=xlUp
    Rows("73:73").Select
    Selection.Delete Shift:=xlUp
    Rows("74:74").Select
    Selection.Delete Shift:=xlUp
    Rows("75:75").Select
    Selection.Delete Shift:=xlUp
    Range("A72").Select
    ActiveCell.FormulaR1C1 = "Type 1 Output, $0.06/lb."
    Range("A73").Select
    ActiveCell.FormulaR1C1 = "Type 2 Output, $0.03/lb."
    Range("A74").Select
    ActiveCell.FormulaR1C1 = "CAAF, $0.03/lb."
    Range("A75").Select
    ActiveCell.FormulaR1C1 = "Cement Kiln feedstock, $0.03/lb"
    Range("A76").Select
    ActiveCell.FormulaR1C1 = "Total Requested ($)"
    Range("A77").Select
    
    'To find the first empty column on the table to insert the total'
    'find last used cell on the row to the right
    Range("A1").Select
    ActiveCell.End(xlToRight).Select
    'move one cell to the right from the last used cell
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "Total"
    
    'Format Tier1_Forecast'
    Columns("A:A").Select
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
    Range("D17").Select
    
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "COMPANY NAME HERE"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Range("A3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Range("A4").Select
    ActiveCell.FormulaR1C1 = _
        "Number of Full Time Equivalent (FTE) Employees in State of California working on carpet recycling"
    
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "2"
    Selection.AutoFill Destination:=Range("A2:A76"), Type:=xlFillSeries
    Range("A2:A76").Select
    Columns("A:A").ColumnWidth = 4
    Range("B52").Select
    ActiveWindow.SmallScroll Down:=-57
    Range("B4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("F4").Select
    Selection.ClearContents
    Rows("77:77").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B79").Select
    ActiveWindow.SmallScroll Down:=-78
    
    '*********************************************************************************'
    'Select the table to insert borders'
    Range("A1:K76").Select
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
    Range("B10").Select
    ActiveWindow.SmallScroll Down:=-27
    
    Range("B71").Select
    ActiveCell.FormulaR1C1 = "Calculations for funding"
    Range("B71").Select
    Range("A71").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("B71").Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("E4").Select
    Selection.ClearContents
    Columns("F:F").Select
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
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B9").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds directly collected by you for this quarter [Do NOT report pounds you are purchasing from other collectors]"
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
    Range("A4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B9:E9").Select
    Selection.Font.Bold = True
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
    Range("B2").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Range("B13").Select
    ActiveCell.FormulaR1C1 = _
        "Carpet directly collected by YOU from California by FIBER type [Do NOT report pounds you are purchasing from other collectors]"
    Range("B13:E13").Select
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
    Range("A13:E13").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B22").Select
    ActiveCell.FormulaR1C1 = _
        "Accounting for total PC Carpet Inputs & Beginning Inventory this quarter"
    Range("B22:E22").Select
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
    Range("A22:E22").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B27").Select
    ActiveCell.FormulaR1C1 = _
        "Accounting for total PC Carpet Outputs & Ending Inventory"
    Range("B27:E27").Select
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
    Range("A27:E27").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("D24").Select
    ActiveWindow.SmallScroll Down:=18
    Range("B40").Select
    ActiveCell.FormulaR1C1 = "Production of Internally Used Whole Carpet"
    Range("B40:E40").Select
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
    Range("A40:E40").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("C43").Select
    ActiveWindow.SmallScroll Down:=18
    Range("B48").Select
    ActiveCell.FormulaR1C1 = _
        "Output and other destinations of post-consumer carpet internally processed this quarter"
    Range("B48:E48").Select
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
    Range("A48:E48").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B52").Select
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
    Selection.Font.Bold = True
    Range("B61").Select
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
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B58").Select
    ActiveCell.FormulaR1C1 = "Total Type 1 Output: SOLD & SHIPPED"
    Range("B58").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
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
    Range("B59").Select
    Selection.Font.Bold = True
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
    
    Range("B71").Select
    ActiveCell.FormulaR1C1 = "Calculations for funding"
    Range("B71:E71").Select
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
    Range("A71:E71").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("D61").Select
    ActiveWindow.SmallScroll Down:=0
    Range("B76").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    
    '*******************************************************'
    'Wrapping the text to fit into column B'
    Range("B10").Select
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
    Range("B11").Select
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
    ActiveWindow.SmallScroll Down:=15
    Range("B31").Select
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
    
End Sub
