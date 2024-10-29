Attribute VB_Name = "Module4"
Private Sub CommandButton1_Click()
 
  Dim s As String, FileName As String, FileNum As Integer
 
  ' Define full pathname of TXT file
  FileName = ThisWorkbook.Path & "\test.txt"
 
  ' Copy range to the clipboard
  Range("A1", Cells(Rows.Count, "A").End(xlUp)).Copy
 
  ' Copy column content to the 's' variable via clipboard
  With New DataObject
     .GetFromClipboard
     s = .GetText
  End With
  Application.CutCopyMode = False
 
  ' Write s to TXT file
  FileNum = FreeFile
  If Len(Dir(FileName)) > 0 Then Kill FileName
  Open FileName For Binary Access Write As FileNum
  Put FileNum, , s
  Close FileNum
 
End Sub



