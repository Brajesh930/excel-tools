Attribute VB_Name = "Module4"
Sub removeFirstLineOnly()
Attribute removeFirstLineOnly.VB_ProcData.VB_Invoke_Func = "E\n14"
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
Dim str As String
Dim selectedCell As Range
Dim arrSplit As Variant
On Error Resume Next
For Each selectedCell In Selection.Cells
str = selectedCell.Value
selectedCell = Replace(selectedCell.Value, Chr(10) + Chr(10), Chr(10))
arrSplit = Split(selectedCell.Value, Chr(41) + Chr(10))
selectedCell = ""
For i = 1 To UBound(arrSplit) - 1
selectedCell = selectedCell.Value + arrSplit(i) + Chr(41) + Chr(10)
If selectedCell = "" Then
selectedCell = str
End If
Next
selectedCell = selectedCell.Value + arrSplit(UBound(arrSplit))
Next
ending:
Application.ScreenUpdating = True
End Sub

