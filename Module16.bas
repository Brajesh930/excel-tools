Attribute VB_Name = "Module16"
Sub ClaimCampwords()

    Dim str1 As String, str2 As String
    Dim arr1 As Variant, arr2 As Variant
    Dim i As Integer, j As Integer
    Dim start As Integer, len1 As Integer
    Dim match As Boolean
    Dim endpos As Integer
    Dim rng As Range, element As Range
Dim element1 As Range
Dim bEntireColumn As Boolean
Dim bEntireRow As Boolean
Dim bEntireColumn1 As Boolean
Dim bEntireRow1 As Boolean
endpos = 0
Dim endposi As Integer
endposi = 0
Dim startPos As Integer
startPos = 0
Dim startposi As Integer
startposi = 0
Dim counter As Integer
counter = 0
Dim counteri As Integer
counteri = 0
Dim countc As Integer
countc = 0
Dim countd As Integer
countd = 0
len1 = 0
Dim intResult As Long
match = False
'reult = Application.Dialogs(xlDialogEditColor).Show(40, 100, 100, 200)
'gets the color selected by the user
intResult = ThisWorkbook.Colors(45)
On Error Resume Next
    Set element1 = Application.InputBox( _
      Title:="Claim Element Input Box", _
      Prompt:="Select a cell of claim element for 1", _
      Type:=8)
  On Error GoTo jump1
With element1
bEntireColumn = .Address = .EntireColumn.Address
bEntireRow = .Address = .EntireRow.Address
End With
If bEntireColumn Then GoTo jump1
If bEntireRow Then GoTo jump1

On Error Resume Next
    Set rng = Application.InputBox( _
      Title:="Multi Claim Elements", _
      Prompt:="Select a cell of claim element for 2", _
      Type:=8)
  On Error GoTo jump1
With rng
bEntireColumn1 = .Address = .EntireColumn.Address
bEntireRow1 = .Address = .EntireRow.Address
End With
If bEntireColumn1 Then GoTo jump1
If bEntireRow1 Then GoTo jump1

  
    'get the strings from cells A1 and B1
    str1 = rng.Text
    str1 = Replace(str1, Chr(10), " ")
    str2 = element1.Text
    str2 = Replace(str2, Chr(10), " ")
    'split strings into arrays of words
    arr1 = Split(str1, " ")
    arr2 = Split(str2, " ")

    'exit if either array has less than 3 words
    If UBound(arr1) < 2 Or UBound(arr2) < 2 Then
        MsgBox "Both cells must contain at least 3 words."
        Exit Sub
    End If
    match = False
    'loop through first array
    For i = 0 To UBound(arr1) - 3
        'initialize match flag
        startposi = startposi + Len(arr1(i)) + 1
        startPos = 0
        'loop through second array
        For j = 0 To UBound(arr2) - 3
            'compare subarrays of 3 words
            startPos = startPos + Len(arr2(j)) + 1
            
            If arr1(i) = arr2(j) And arr1(i + 1) = arr2(j + 1) And arr1(i + 2) = arr2(j + 2) Then
                'set match flag
                match = True
                'store start and length of match
                
                endpos = startPos + Len(arr2(j)) + Len(arr2(j + 1)) + Len(arr2(j + 2)) + 3
                len1 = 3
                'loop to check for longer match
                Do While i + len1 < UBound(arr1) And arr1(i + len1) = arr2(j + len1) And j + len1 < UBound(arr2)
                    endpos = endpos + Len(arr2(j + len1)) + 1
                    len1 = len1 + 1
                Loop
                '
                'highlight matched words in second string
                counter = endpos - startPos
                'If element1.Characters(startPos - Len(arr2(j)), counter).Font.Color = RGB(0, 0, 0) Then
                
                element1.Characters(startPos - Len(arr2(j)), counter).Font.Color = intResult
                rng.Characters(startposi - Len(arr1(i)), counter).Font.Color = intResult
                'End If
                startposi = startposi + (counter - Len(arr1(i)) - 1)
                'exit inner loop
                i = i + (len1 - 1)
                
                Exit For
            End If
        Next j
        intResult = intResult + 4500
        
    Next i
jump1:
    If match = False Then
        MsgBox "No matching sequence of words found. or wrong selection of cell"
    End If
End Sub



