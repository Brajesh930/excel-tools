Attribute VB_Name = "Module7"
Public Sub FitIndividualPic()
Dim Pic As Shapes
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

