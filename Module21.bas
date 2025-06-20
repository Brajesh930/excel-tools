Attribute VB_Name = "Module21"
Sub Highlighterforexcel()
    Dim tableRange As Range
    Set tableRange = Application.InputBox("Select the range of the table:", "Table Range", Type:=8)
    Dim rng As Range
    Set rng = Application.InputBox("Select the range of the cells to be highlighted:", "Cell Range", Type:=8)
    Dim numberOfColumns As Integer
    numberOfColumns = tableRange.Columns.count
    Dim i As Integer
    Dim lastColor As Integer 'variable to store the color of the last word added
    For Each cell In rng
        Dim foundWords() As Variant
        ReDim foundWords(0 To 1, 0)
        Dim startPos As Integer
        startPos = 0
        For i = 1 To numberOfColumns
            For Each w In tableRange.Columns(i).Cells
                If Not IsEmpty(w) Then
                    Do While InStr(startPos + 1, cell.Value, w.Value) > 0
                        cell.Characters(InStr(startPos + 1, cell.Value, w.Value), Len(w.Value)).Font.ColorIndex = w.Font.ColorIndex
                        ReDim Preserve foundWords(1, UBound(foundWords, 2) + 1)
                        foundWords(0, UBound(foundWords, 2)) = w.Value
                        foundWords(1, UBound(foundWords, 2)) = tableRange.Cells(1, i).Font.ColorIndex
                        startPos = InStr(startPos + 1, cell.Value, w.Value)
                    Loop
                    startPos = 0
                End If
            Next w
        Next i
        Dim foundWordsString As String
        foundWordsString = "" 'clear the foundWordsString variable before starting the loop
        For j = 0 To UBound(foundWords, 2)
            If j = 0 Then
                lastColor = foundWords(1, j) 'initialize lastColor with the color of the first word
            Else
                If foundWords(1, j) <> lastColor Then 'if the color of the current word is different from the last word
                    foundWordsString = foundWordsString & vbNewLine 'add a new line
                    lastColor = foundWords(1, j) 'update lastColor with the color of the current word
                End If
            End If
            foundWordsString = foundWordsString & " " & foundWords(0, j)
        Next j
        cell.Offset(0, 1).Value = foundWordsString
        For j = 0 To UBound(foundWords, 2)
            cell.Offset(0, 1).Characters(InStr(1, cell.Offset(0, 1).Value, foundWords(0, j)), Len(foundWords(0, j))).Font.ColorIndex = foundWords(1, j)
        Next j
    Next cell
End Sub


