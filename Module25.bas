Attribute VB_Name = "Module25"
Sub HyperLinksUSPTOEspacent()

Dim cell As Range
    Dim patentNum As String
    Dim cleanedPatentNum As String
    Dim hyperlinkURL As String
    Dim screenTipText As String
    Dim proceed As VbMsgBoxResult

    For Each cell In Selection
        If IsEmpty(cell.Value) Then
            proceed = MsgBox("Empty cell found at " & cell.Address & ". Do you want to continue?", vbYesNo + vbExclamation)
            If proceed = vbNo Then Exit Sub
        Else
            patentNum = Trim(cell.Value)

            ' Check if it's a US patent
            If UCase(Left(patentNum, 2)) = "US" Then
                ' Remove country code (US) and kind code (last 2 chars like A1, B2)
                cleanedPatentNum = Mid(patentNum, 3)
                If cleanedPatentNum Like "*[A-Z][0-9]" Then
                    cleanedPatentNum = Left(cleanedPatentNum, Len(cleanedPatentNum) - 2)
                End If

                hyperlinkURL = "https://ppubs.uspto.gov/pubwebapp/authorize.html?redirect=patents/html/" & cleanedPatentNum
                screenTipText = "Link To USPTO"
            Else
                ' Non-US patent - use full number
                hyperlinkURL = "https://worldwide.espacenet.com/patent/search?q=pn%3D" & patentNum
                screenTipText = "Link To Espacenet"
            End If

            ' Add hyperlink with ScreenTip
            cell.Hyperlinks.Add Anchor:=cell, Address:=hyperlinkURL, _
                TextToDisplay:=patentNum, ScreenTip:=screenTipText
        End If
    Next cell

    MsgBox "Hyperlinks created", vbInformation
End Sub



