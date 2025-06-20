Attribute VB_Name = "Module22"
Sub CreateAndOpenHyperlinksInNextColumn()
    Dim selectedRange As Range
    Dim cell As Range
    Dim targetCell As Range
    Dim hyperlinkAddress As String
    ' Prompt the user to select a range of cells with plain text
    On Error Resume Next
    Set selectedRange = Application.InputBox("Please select a range of cells with plain text.", Type:=8)
    On Error GoTo 0
    ' Check if a range was selected
    If Not selectedRange Is Nothing Then
        ' Loop through each cell in the selected range
        For Each cell In selectedRange
            ' Check if the cell is not empty
            If Not IsEmpty(cell.Value) Then
                ' Get the corresponding cell in the next column
                Set targetCell = cell.Offset(0, 1)
                ' Create a hyperlink in the next cell
                hyperlinkAddress = "https://www.google.com/search?q=" & cell.Value ' Modify as needed
                targetCell.Hyperlinks.Add _
                    Anchor:=targetCell, _
                    Address:=hyperlinkAddress, _
                    TextToDisplay:=cell.Value
                ' Open the hyperlink in the next cell
                targetCell.Hyperlinks(1).Follow
            End If
        Next cell
    Else
        MsgBox "No range of cells was selected."
    End If
End Sub

