Attribute VB_Name = "Module1"
Sub formatcell()
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
Dim str As String
Dim selectedCell As Range
On Error Resume Next
For Each selectedCell In Selection.Cells
selectedCell.WrapText = True
selectedCell.HorizontalAlignment = xlCenter
selectedCell.VerticalAlignment = xlCenter
selectedCell = selectedCell.Value
If IsDate(selectedCell) Then
selectedCell.NumberFormat = "mmmm d, yyyy;@"
End If
Next
ending:
Application.ScreenUpdating = True
End Sub
