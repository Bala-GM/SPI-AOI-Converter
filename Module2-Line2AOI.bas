Attribute VB_Name = "Module2"
Sub Line2AOI()
Attribute Line2AOI.VB_ProcData.VB_Invoke_Func = "W\n14"
'
' Line2AOI Macro
'
' Keyboard Shortcut: Ctrl+Shift+W
'
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AC$2000").AutoFilter Field:=10, Criteria1:="Yes"
    Range("A2:S2000").Select
    Selection.ClearContents
    ActiveSheet.Range("$A$1:$AC$2000").AutoFilter Field:=10, Criteria1:="<>"
    Range("G13").Select
    ActiveSheet.Range("$A$1:$AC$2000").AutoFilter Field:=1, Criteria1:="1"
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:R").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:F").Select
    ActiveWindow.SmallScroll Down:=2000
    Selection.Copy
    ActiveWindow.SmallScroll Down:=-2000
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Range("K10").Select
    ActiveWindow.SmallScroll Down:=-2000
    Rows("1:1").Select
    MsgBox "AOI Data Conversion Complete Thank You"
End Sub
