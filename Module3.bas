Attribute VB_Name = "Module3"
Sub removeFirstLine()
Attribute removeFirstLine.VB_ProcData.VB_Invoke_Func = "R\n14"
Dim bEntireColumn As Boolean
Dim bEntireRow As Boolean
With Selection
bEntireColumn = .Address = .EntireColumn.Address
bEntireRow = .Address = .EntireRow.Address
End With
If bEntireColumn Then
MsgBox "Entire column(s) selected ---select again"
GoTo ending
End If
If bEntireRow Then
MsgBox "Entire row(s) selected---select again "
GoTo ending
End If
Application.ScreenUpdating = False
Dim str As String
Dim selectedCell As Range
Dim arrSplit As Variant
On Error Resume Next
For Each selectedCell In Selection.Cells
arrSplit = Split(selectedCell.Value, Chr(10))
For i = 1 To UBound(arrSplit)
str = str + arrSplit(i) + Chr(10)
Next
selectedCell = str
str = ""
Next
ending:
Application.ScreenUpdating = True
End Sub
