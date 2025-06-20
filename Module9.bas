Attribute VB_Name = "Module9"
Sub ClaimCamp()
Dim rng As Range, element As Range
Dim element1 As Range
Dim bEntireColumn As Boolean
Dim bEntireRow As Boolean
Dim bEntireColumn1 As Boolean
Dim bEntireRow1 As Boolean
k = 0
Dim intResult As Long
'displays the color dialog
On Error Resume Next
    Set element = Application.InputBox( _
      Title:="Claim Element Input Box", _
      Prompt:="Select a cell of claim element", _
      Type:=8)
  On Error GoTo 0
  
With element
bEntireColumn = .Address = .EntireColumn.Address
bEntireRow = .Address = .EntireRow.Address
End With
If bEntireColumn Then GoTo jump2
If bEntireRow Then GoTo jump2
  reult = Application.Dialogs(xlDialogEditColor).Show(40, 100, 100, 200)
'gets the color selected by the user
intResult = ThisWorkbook.Colors(40)
   On Error Resume Next
    Set rng = Application.InputBox( _
      Title:="Multi Claim Elements", _
      Prompt:="Select Range to hilight", _
      Type:=8)
  On Error GoTo 0
With rng
bEntireColumn1 = .Address = .EntireColumn.Address
bEntireRow1 = .Address = .EntireRow.Address
End With
If bEntireColumn1 Then GoTo jump2
If bEntireRow1 Then GoTo jump2
  For Each row1 In element
  If row1.Value2 = Null Then GoTo jump2
 text1 = row1.Value2
 For Each row In rng
 If row.Value2 = Null Then GoTo jump1
k = InStr(1, row.Value2, text1, vbTextCompare)
If k > 0 Then
row.Characters(k, Len(text1)).Font.Color = intResult
End If
Next row
jump1:
intResult = intResult + 4500
Next row1
jump2:
End Sub


