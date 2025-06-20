Attribute VB_Name = "Module5"
Sub FixedSize154()
Application.ScreenUpdating = False
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
Dim selectedCell As Range
On Error Resume Next
For Each selectedCell In Selection.Cells
selectedCell.RowHeight = 156
selectedCell.HorizontalAlignment = xlCenter
selectedCell.VerticalAlignment = xlCenter
Next
ending:
Application.ScreenUpdating = True
End Sub
