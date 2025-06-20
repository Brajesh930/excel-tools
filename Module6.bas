Attribute VB_Name = "Module6"
Public Sub FitMultipleSelectedPics()
Dim Pic As Object
Dim answer As Integer
Dim countpic As Integer
countpic = 0
donepic = 1
For Each Pic In Selection.ShapeRange
            countpic = countpic + 1
        Next Pic
If countpic > 10 Then
againfit:
ShapeSelection1 (donepic)
For Each Pic In Selection.ShapeRange
            FitIndividualPic Pic
            donepic = donepic + 1
        Next Pic
If donepic < countpic Then GoTo againfit
Else
For Each Pic In Selection.ShapeRange
            FitIndividualPic Pic
        Next Pic
        End If
End Sub

Public Sub FitIndividualPic(Pic As Object)
    Dim PicWtoHRatio As Single
    Dim CellWtoHRatio As Single
    Dim Gap As Single
    Gap = 10
    With Pic
        PicWtoHRatio = (.Width / (.Height + 1))
    End With

    With Pic.TopLeftCell
        CellWtoHRatio = .Width / .RowHeight
     End With

    Select Case PicWtoHRatio / CellWtoHRatio
        Case Is > 1
            With Pic
                .Width = .TopLeftCell.Width - Gap
                .Height = .Width / PicWtoHRatio - Gap
            End With
        Case Else

            With Pic
                .Height = .TopLeftCell.RowHeight - Gap
                .Width = .Height * PicWtoHRatio - Gap
            End With
    End Select

    With Pic
        .Top = .TopLeftCell.Top + Gap
        .Left = .TopLeftCell.Left + Gap
    End With
End Sub
Public Sub ShapeSelection1(starting As Integer)
Dim Sh As Shape
Dim selectedOne As Boolean
Dim count As Integer
count = 0
ending = starting + 10
On Error Resume Next
    With ActiveSheet
       For Each Sh In .Shapes
       If count < starting Then GoTo end1
           If Not Application.Intersect(Sh.TopLeftCell, .Range(Selection.Address)) Is Nothing Then
              If selectedOne = False Then
                  Sh.Select
                  selectedOne = True
               Else
                  Sh.Select (False)
               End If
           End If
end1:
           If count > ending Then
           Exit For
           Else
           count = count + 1
           End If
        Next Sh
    End With
End Sub
