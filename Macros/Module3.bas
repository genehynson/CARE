Attribute VB_Name = "Module3"
Sub Tier1ActualMacro()
Attribute Tier1ActualMacro.VB_Description = "This MACRO formats the tier 1 Actual spreddsheet."
Attribute Tier1ActualMacro.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' Tier1ActualMacro Macro
' This MACRO formats the tier 1 Actual spreddsheet.
'
' Keyboard Shortcut: Ctrl+e
'
    Application.Run "PERSONAL.XLSB!GenerateReports"
    
    Sheets("Tier1_Actual").Select
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Cut
    Rows("3:3").Select
    Selection.Insert Shift:=xlDown
    Selection.Cut
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "COMPANY NAME HERE"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "CONFIDENTIAL"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "AB 2398 Monthly Rolling Forecast"
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A4").Select
    ActiveCell.FormulaR1C1 = _
        "Number of Full Time Equivalent (FTE) Employees in State of California working on carpet recycling"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = _
        "Number of CA FTE Employees at beginning of this quarter"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "Number of FTE CA Jobs Lost this quarter"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "Number of FTE CA Jobs gained this quarter"
    Range("A8").Select
    ActiveCell.FormulaR1C1 = "Number of FTE CA Employees at end of this quarter"
    Rows("9:9").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A9").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds directly collected by you for this quarter [Do NOT report pounds you are purchasing from other collectors]"
    Range("A10").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds directly collected by you from California this quarter"
    Range("A11").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds directly collected by you from OUTSIDE California for this quarter"
    Rows("12:12").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A12").Select
    ActiveCell.FormulaR1C1 = "TOTAL Post-consumer carpet pounds"
    Rows("13:13").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A13").Select
    ActiveCell.FormulaR1C1 = _
        "Carpet directly collected by YOU from California by FIBER type [Do NOT report pounds you are purchasing from other collectors]"
    Range("A14").Select
    ActiveCell.FormulaR1C1 = "Nylon 6"
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "Nylon6, 6"
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "Nylon 6, 6"
    Range("A16").Select
    ActiveCell.FormulaR1C1 = "Polypropylene"
    Range("A17").Select
    ActiveCell.FormulaR1C1 = "PET"
    Range("A18").Select
    ActiveCell.FormulaR1C1 = "Wool"
    Range("A19").Select
    ActiveCell.FormulaR1C1 = "Other/Mixed Fibers"
    Rows("20:20").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A20").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("21:21").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A21").Select
    ActiveCell.FormulaR1C1 = "Line 20 must equal Line 10"
    Rows("22:22").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A22").Select
    ActiveCell.FormulaR1C1 = _
        "Accounting for total PC Carpet Inputs & Beginning Inventory this quarter"
    Range("A23").Select
    ActiveCell.FormulaR1C1 = _
        "Beginning Inventory of Whole Carpet from CA at start of quarter (should equal prior quarter ending inventory)."
    Range("A24").Select
    ActiveCell.FormulaR1C1 = "Whole Carpet Collected from California (Row 10)"
    Range("A25").Select
    ActiveCell.FormulaR1C1 = "Whole Carpet from CA received from other collectors"
    Rows("26:26").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A26").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("27:27").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A27").Select
    ActiveCell.FormulaR1C1 = _
        "Accounting for total PC Carpet Outputs & Ending Inventory"
    Range("A28").Select
    ActiveCell.FormulaR1C1 = "Re-Used"
    Range("A29").Select
    ActiveCell.FormulaR1C1 = "Internally Used Whole Carpet this quarter"
    Range("A30").Select
    ActiveCell.FormulaR1C1 = _
        "Whole carpet shipped to US customers OUTSIDE California"
    Range("A31").Select
    ActiveCell.FormulaR1C1 = _
        "Whole carpet shipped to customers outside the United States"
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
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A38").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Range("A39").Select
    ActiveCell.FormulaR1C1 = "Line 38 must equal Line 26"
    Rows("40:40").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A40").Select
    ActiveCell.FormulaR1C1 = "Production of Internally Used Whole Carpet"
    Range("A41").Select
    ActiveCell.FormulaR1C1 = "Internally Used Whole Carpet this quarter"
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
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A46").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Range("A47").Select
    ActiveCell.FormulaR1C1 = "Line 46 must equal Line 41"
    Range("A48").Select
    ActiveCell.FormulaR1C1 = _
        "Output and other destinations of post-consumer carpet internally processed this quarter"
    Range("A49").Select
    
    ActiveCell.FormulaR1C1 = _
        "Beginning Inventory of Processed Goods from prior quarter"
    Range("A50").Select
    ActiveCell.FormulaR1C1 = "Processed"
    Rows("51:51").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A51").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Range("A52").Select
    ActiveCell.FormulaR1C1 = "Type 1 Outputs"
    Range("A53").Select
    ActiveCell.FormulaR1C1 = "Fiiber"
    Range("A54").Select
    ActiveCell.FormulaR1C1 = "DePoly or Chemical Component"
    Range("A55").Select
    ActiveCell.FormulaR1C1 = "Shredded Carpet tile used for tile backing"
    Range("A56").Select
    ActiveCell.FormulaR1C1 = _
        "Number of Ash Tests run this quarter (min 1 per 1M pounds)"
    Range("A57").Select
    ActiveCell.FormulaR1C1 = _
        "Average Ash Test Results over quarter for Type 1 pounds"
    Range("A58").Select
    ActiveWindow.SmallScroll Down:=12
    Range("A60").Select
    ActiveCell.FormulaR1C1 = "Filler"
    Range("A61").Select
    ActiveCell.FormulaR1C1 = "CAAF_Payout_Per_Lb"
    Range("A61").Select
    ActiveCell.FormulaR1C1 = "CAAF"
    Range("A62").Select
    ActiveCell.FormulaR1C1 = "Cement Kiln Feedstock"
    Range("A63").Select
    ActiveCell.FormulaR1C1 = "Carcass Sold"
    Rows("64:64").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A64").Select
    ActiveCell.FormulaR1C1 = "Total Type 2 Output: SOLD & SHIPPED"
    Range("A65").Select
    ActiveCell.FormulaR1C1 = "Landfilled"
    Range("A66").Select
    ActiveCell.FormulaR1C1 = "WTE"
    Range("A67").Select
    ActiveCell.FormulaR1C1 = "Incinerated"
    Range("A68").Select
    ActiveCell.FormulaR1C1 = "Ending Inventory Processed Goods this quarter"
    Range("A69").Select
    
    Rows("70:70").Select
    Selection.Delete Shift:=xlUp
    Rows("71:71").Select
    Selection.Delete Shift:=xlUp
    Rows("72:72").Select
    Selection.Delete Shift:=xlUp
    Rows("73:73").Select
    Selection.Delete Shift:=xlUp
    Rows("69:69").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A69").Select
    ActiveCell.FormulaR1C1 = "TOTAL Recycled Pounds This Quarter"
    Rows("70:70").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A70").Select
    ActiveCell.FormulaR1C1 = "Line 69 must equal Line 51"
    Rows("71:71").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A71").Select
    ActiveCell.FormulaR1C1 = "Calculations for funding"
    Range("A72").Select
    ActiveCell.FormulaR1C1 = "Type 1 Output, $0.06/lb."
    Range("A73").Select
    ActiveCell.FormulaR1C1 = "Type 2 Output, $0.03/lb."
    Range("A74").Select
    ActiveCell.FormulaR1C1 = "CAAF, $0.03/lb."
    Range("A75").Select
    ActiveCell.FormulaR1C1 = "Cement Kiln Feedstock"
    Range("A75").Select
    ActiveCell.FormulaR1C1 = "Cement Kiln Feedstock, $0.03/lb."
    Range("A76").Select
    ActiveCell.FormulaR1C1 = "Total Requested($)"
    Rows("77:77").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A79").Select
    ActiveCell.FormulaR1C1 = "New 10 CENT GROWTH INCENTIVE CALCULATIONS"
    Range("A81").Select
    ActiveCell.FormulaR1C1 = "Total Type 1 Pounds for Quarter"
    Range("A82").Select
    ActiveCell.FormulaR1C1 = "Target Pounds for growth incentive"
    Range("A83").Select
    ActiveCell.FormulaR1C1 = "Over (Under) Target"
    Range("A84").Select
    ActiveCell.FormulaR1C1 = "Total Growth Incentive Pool"
    Range("A86").Select
    ActiveCell.FormulaR1C1 = "Percent Contribution by Each Processor for Type 1"
    Range("A87").Select
    ActiveCell.FormulaR1C1 = "GROWTH Payout to Each Processor"
    Rows("88:88").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A90").Select
    ActiveCell.FormulaR1C1 = "TOTAL PAYOUTS for QUARTER"
    Rows("92:92").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A92").Select
    ActiveCell.FormulaR1C1 = "CORECTIONS OR ADJUSTMENTS"
    Range("A94").Select
    ActiveCell.FormulaR1C1 = "GRAND TOTAL PAYOUTS"
   
    Range("A58").Select
    ActiveCell.FormulaR1C1 = "Total Type 1 Output: SOLD & SHIPPED"
    Rows("59:59").Select
    Selection.Cut Destination:=Rows("107:107")
    Range("A97").Select
    ActiveWindow.SmallScroll Down:=-45
    Range("A59").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Type 2 Outputs"
    
    'To find the first empty column on the table to insert the total'
    'find last used cell on the row to the right
    Range("A1").Select
    ActiveCell.End(xlToRight).Select
    'move one cell to the right from the last used cell
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "Total"
    
    'Adding colors, Borders, etc.'
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B4").Select
    Columns("B:B").EntireColumn.AutoFit
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "2"
    Selection.AutoFill Destination:=Range("A2:A76"), Type:=xlFillSeries
    Range("A2:A76").Select
    Columns("A:A").ColumnWidth = 4
    Columns("B:B").Select
    Range("B44").Activate
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
    ActiveWindow.SmallScroll Down:=-57
    Selection.ColumnWidth = 56.22
    Selection.Rows.AutoFit
    ActiveWindow.SmallScroll Down:=-48
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
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("B1").Select
    Selection.Font.Bold = True
    Range("B3").Select
    Selection.Font.Bold = True
    Range("B4:E4").Select
    Range("B4").Select
    Range(Selection, Selection.End(xlToRight)).Select
   
    Range("B4").Select
    ActiveWindow.SmallScroll Down:=57
    Range("C94").Select
    ActiveCell.FormulaR1C1 = "=+R[-4]C+R[-2]C"
    Range("D94").Select
    ActiveCell.FormulaR1C1 = "=+R[-4]C+R[-2]C"
    Range("D95").Select
    ActiveWindow.SmallScroll Down:=-105
    Range("C1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=72
    Range("B85").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("C85").Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.Paste
    Range("E94").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=+R[-4]C+R[-2]C"
    Range("E95").Select
    ActiveWindow.SmallScroll Down:=-99
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3])"
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("F5").Select
    Range("F5").Select
   
    Range("C94").Select
    ActiveWindow.SmallScroll Down:=-90
    Range("B4").Select
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
    ActiveWindow.SmallScroll Down:=9
    Range("B22:D22").Select
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
    ActiveWindow.SmallScroll Down:=-12
    Range("B1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Rows("11:11").Select
    Rows("11:11").EntireRow.AutoFit
    Rows("11:11").EntireRow.AutoFit
    Rows("10:10").EntireRow.AutoFit
    Rows("10:10").Select
    Selection.RowHeight = 25.2
    Rows("10:10").Select
    Selection.Rows.AutoFit
    Selection.RowHeight = 27.6
    Range("B15").Select
    ActiveWindow.SmallScroll Down:=-21
    
    'Select the table to insert columns'
    Range("A1:K76").Select
    Range("B1").Activate
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
    Range("B22:D22").Select
    ActiveWindow.SmallScroll Down:=60
    Range("B79:F88").Select
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
    ActiveWindow.SmallScroll Down:=3
    Range("C90").Select
    ActiveWindow.SmallScroll Down:=15
    Range("B90:F90").Select
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
    Range("B92:F92").Select
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
    Range("B94:F94").Select
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
    ActiveWindow.SmallScroll Down:=-57
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
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("A71:E71").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    ActiveWindow.SmallScroll Down:=-9
    Selection.Font.Bold = True
    Range("B76").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Range("B64").Select
    Selection.Font.Bold = True
    ActiveWindow.SmallScroll Down:=-18
    Range("A48:E48").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    ActiveWindow.SmallScroll Down:=-27
    Range("B22:D22").Select
    ActiveWindow.SmallScroll Down:=15
    Range("A40:E40").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B40").Select
    Range(Selection, Selection.End(xlToRight)).Select
    
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
    ActiveWindow.SmallScroll Down:=-39
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Range("B4:E4").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B9:E9").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B13:E13").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("B9:E9").Select
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
    Range("B13:E13").Select
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
    Range("B22:E22").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
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
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Range("F5").Select
    ActiveWindow.SmallScroll Down:=-21
    Range("F5").Select
    Selection.ClearContents
    Range("F6").Select
    Selection.ClearContents
    
    '*******************************************************'
    'Wrapping the text to fit on column B'
    Columns("B:B").ColumnWidth = 46.11
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
    ActiveWindow.SmallScroll Down:=12
    Range("B23").Select
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
    Range("B30").Select
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
