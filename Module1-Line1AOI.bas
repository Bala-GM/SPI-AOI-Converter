Attribute VB_Name = "Module1"
Sub Line1AOI()
Attribute Line1AOI.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' Line1AOI Macro
'
' Keyboard Shortcut: Ctrl+Shift+A
'
    Rows("1:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:K").Select
    Range("K1").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "p"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "w"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "f"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "f"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "d"
    Range("A5:A2000").Select
    Selection.FillDown
    ActiveWindow.SmallScroll Down:=-2100
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "=-RC[13]*1000"
    Range("B6").Select
    Selection.AutoFill Destination:=Range("B6:B2000"), Type:=xlFillDefault
    Range("B6:B2000").Select
    ActiveWindow.SmallScroll Down:=-2100
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "=RC[11]*1000"
    Range("C6").Select
    Selection.AutoFill Destination:=Range("C6:C2000"), Type:=xlFillDefault
    Range("C6:C2000").Select
    ActiveWindow.SmallScroll Down:=-2100
    Range("D6").Select
    ActiveCell.FormulaR1C1 = "=RC[8]&"":""&RC[9]"
    Range("D6").Select
    Selection.AutoFill Destination:=Range("D6:D2000"), Type:=xlFillDefault
    Range("D6:D2000").Select
    ActiveWindow.SmallScroll Down:=-2100
    Range("E6").Select
    ActiveCell.FormulaR1C1 = "n0000"
    Range("E6:E2000").Select
    Selection.FillDown
    ActiveWindow.SmallScroll Down:=-2100
    Columns("Q:Q").Select
    Selection.Copy
    Columns("F:F").Select
    ActiveSheet.Paste
    Columns("R:R").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("G:G").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    Columns("AD:AD").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Columns("H:H").Select
    ActiveSheet.Paste
    Range("I6").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[3]&"":""&RC[4]"
    Range("I6").Select
    Selection.AutoFill Destination:=Range("I6:I2000"), Type:=xlFillDefault
    Range("I6:I2000").Select
    ActiveWindow.SmallScroll Down:=-2100
    Range("J6").Select
    ActiveCell.FormulaR1C1 = "shape"
    Range("J6:J2000").Select
    Selection.FillDown
    ActiveWindow.SmallScroll Down:=-2100
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("B2").Select
    MsgBox "AOI Data Conversion Complete Thank You"
End Sub
