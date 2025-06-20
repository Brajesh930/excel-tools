Attribute VB_Name = "Module2"
Sub TextToDisplayForHyperlink()
Application.ScreenUpdating = False
Dim rng As Range, cell As Range
Set rng = Selection
For Each cell In rng
If IsEmpty(ActiveCell) = True Then
Dim answer As Integer
answer = MsgBox("Cell is empty in  " & ActiveCell.row & "  row and " & ActiveCell.Column & "  column", vbQuestion + vbYesNo + vbDefaultButton2, "Do you want to exit")
    If answer = vbYes Then
     Exit For
     Else: GoTo jump
     End If
    End If
ActiveCell.Hyperlinks.Add Anchor:=ActiveCell, Address:="https://worldwide.espacenet.com/patent/search/" + ActiveCell.Text + "?q=" + ActiveCell.Text, ScreenTip:="Link to Espacenet", TextToDisplay:=ActiveCell.Text
jump:
ActiveCell.Offset(1, 0).Select
Next cell
Application.ScreenUpdating = True
End Sub

